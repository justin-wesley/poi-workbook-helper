package com.wesleyhome.poi.api.report;

import org.apache.commons.lang3.reflect.FieldUtils;

import java.lang.reflect.Field;
import java.util.function.Function;

public class FieldColumnConfiguration<T> extends AbstractAnnotatedColumnConfiguration<T, Field> {

    public FieldColumnConfiguration(Field field, ReportConfiguration<T> reportConfiguration) {
        super(field, reportConfiguration);
    }

    @Override
    protected Function<T, Object> getAccessor(Field annotatedElement) {
        return objectValue -> {
            try {
                //noinspection unchecked
                return FieldUtils.readField(annotatedElement, objectValue, true);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        };
    }

    @Override
    protected Function<Field, String> getPropertyNameFunction() {
        return Field::getName;
    }

}
