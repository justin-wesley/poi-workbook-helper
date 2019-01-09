package com.wesleyhome.poi.api.report;

import org.apache.commons.lang3.reflect.MethodUtils;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.function.Function;
import java.util.stream.Stream;

public class MethodColumnConfiguration<T> extends AbstractAnnotatedColumnConfiguration<T, Method> {

    public MethodColumnConfiguration(Method annotatedElement) {
        super(annotatedElement);
    }

    @Override
    protected Function<T, Object> getAccessor(Method annotatedElement) {
        return objectValue -> {
            try {
                //noinspection unchecked
                return MethodUtils.getAccessibleMethod(annotatedElement).invoke(objectValue);
            } catch (IllegalAccessException | InvocationTargetException e) {
                throw new RuntimeException(e);
            }
        };
    }

    @Override
    protected Function<Method, String> getFieldNameFunction() {
        return method -> {
            try {
                return Stream.of(Introspector.getBeanInfo(method.getDeclaringClass()).getPropertyDescriptors())
                    .filter(pd -> pd.getReadMethod().equals(method))
                    .findFirst()
                    .orElseThrow(IllegalArgumentException::new)
                    .getName();
            } catch (IntrospectionException e) {
                throw new RuntimeException(e);
            }
        };
    }
}
