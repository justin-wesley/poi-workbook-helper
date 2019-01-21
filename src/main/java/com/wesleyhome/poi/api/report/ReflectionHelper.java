package com.wesleyhome.poi.api.report;

import java.lang.annotation.Annotation;
import java.lang.reflect.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Stream;

import static java.util.stream.Collectors.*;

final class ReflectionHelper {

    public static <A extends Annotation, M extends Member & AnnotatedElement> List<M> getAnnotatedMembers(Class<?> annotatedClass, Class<A> annotationClass) {
        Predicate<Method> get = m -> m.getName().startsWith("get") || m.getReturnType().equals(Boolean.TYPE) && m.getName().startsWith("is");
        List<Field> fields = getFields(annotatedClass, null, annotatedPredicate(annotationClass));
        List<Method> methods = getMethods(annotatedClass, null, get.and(annotatedPredicate(annotationClass)));
        Stream<M> fieldStream = fields.stream().map(f->(M)f);
        Stream<M> methodStream = methods.stream().map(f->(M)f);
        Stream<M> concat = Stream.concat(fieldStream, methodStream);
        return concat
            .collect(
            collectingAndThen(
                collectingAndThen(
                    toMap(Member::getName, m -> (M) m, (m1, m2) -> m2 instanceof Field ? m2 : m1),
                    Map::values
                ),
                ArrayList::new
            )
        );
    }

    public static List<Method> getAllMethods(Class<?> startClass, Class<?> traverseToClass) {
        return getMethods(startClass, traverseToClass, truePredicate());
    }

    public static List<Method> getMethods(Class<?> startClass, Class<?> traverseToClass, Predicate<Method> methodPredicate) {
        return getMembers(startClass, traverseToClass, Class::getDeclaredMethods, methodPredicate);
    }

    public static List<Field> getAllFields(Class<?> startClass, Class<?> traverseToClass) {
        return getFields(startClass, traverseToClass, truePredicate());
    }

    public static List<Field> getFields(Class<?> startClass, Class<?> traverseToClass, Predicate<Field> methodPredicate) {
        return getMembers(startClass, traverseToClass, Class::getDeclaredFields, methodPredicate);
    }

    private static <M extends AccessibleObject> List<M> getMembers(Class<?> holdingClass, Class<?> traverseToClass, Function<Class<?>, M[]> memberSupplierFunction, Predicate<M> filterPredicate) {
        List<M> members = Stream.of(memberSupplierFunction.apply(holdingClass)).filter(filterPredicate).collect(toList());
        if (holdingClass.equals(traverseToClass)) {
            return members;
        }
        Class<?> superclass = holdingClass.getSuperclass();
        if (superclass == null || superclass.equals(Object.class)) {
            return members;
        }
        return Stream.concat(
            members.stream(),
            getMembers(superclass, traverseToClass, memberSupplierFunction, filterPredicate).stream()
        ).collect(toList());
    }

    private static <T extends Member> Predicate<T> truePredicate() {
        return m -> true;
    }

    private static <T extends Member & AnnotatedElement, A extends Annotation> Predicate<T> annotatedPredicate(Class<A> annotation) {
        return m -> m.isAnnotationPresent(annotation);
    }
}
