package com.wesleyhome.poi.api.report;

import com.google.common.base.CaseFormat;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.WordUtils;

import java.lang.reflect.AnnotatedElement;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Member;
import java.util.function.Function;

import static org.apache.commons.lang3.reflect.ConstructorUtils.invokeConstructor;

public abstract class AbstractAnnotatedColumnConfiguration<T, M extends Member & AnnotatedElement> implements ColumnConfiguration<T> {
    private final ReportColumn reportColumn;
    private final M annotatedElement;
    private String columnName;
    private final SafeLazyInitializer<Function<Object, Object>> typeTransformerInitializer;
    private final SafeLazyInitializer<Function<T, Object>> accessorInitializer;
    private final SafeLazyInitializer<String> fieldNameInitializer;
    private final SafeLazyInitializer<String> defaultHeaderInitializer;

    public AbstractAnnotatedColumnConfiguration(M annotatedElement) {
        this.reportColumn = annotatedElement.getAnnotation(ReportColumn.class);
        this.annotatedElement = annotatedElement;
        this.columnName = reportColumn.columnIdentifier();
        this.typeTransformerInitializer = initializer(m->{
            Class<? extends Function<?, ?>> formatter = reportColumn.formatter();
            try {
                //noinspection unchecked
                return (Function<Object, Object>) invokeConstructor(formatter);
            } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException | InstantiationException e) {
                throw new RuntimeException(e);
            }
        }) ;

        this.accessorInitializer = initializer(m->getAccessor(annotatedElement));
        this.fieldNameInitializer = initializer(m->getFieldNameFunction().apply(annotatedElement));
        this.defaultHeaderInitializer = initializer(m->getDefaultColumnHeaderNameFunction().apply(annotatedElement));
    }

    public <TY> SafeLazyInitializer<TY> initializer(Function<M, TY> initializer) {
        return new SafeLazyInitializer<TY>() {
            @Override
            protected TY safeInitialize() {
                try {
                    return initializer.apply(annotatedElement);
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
        };
    }

    @Override
    public ColumnType getColumnType() {
        return this.reportColumn.columnType();
    }

    @Override
    public String getColumnHeader() {
        return ReportColumn.NULL.equals(reportColumn.columnHeader()) ? StringUtils.capitalize(getDefaultHeaderName()) : reportColumn.columnHeader();
    }


    public Function<Object, Object> getTypeTransformer() {
        return typeTransformerInitializer.get();
    }

    @Override
    public final Function<T, Object> getAccessor() {
        return accessorInitializer.get();
    }

    protected String getDefaultHeaderName(){
        return defaultHeaderInitializer.get();
    }

    protected Function<M, String> getDefaultColumnHeaderNameFunction() {
        return field-> WordUtils.capitalize(CaseFormat.LOWER_CAMEL.to(
            CaseFormat.UPPER_UNDERSCORE,
            fieldNameInitializer.get()
        ).replace("_", " "));
    }

    @Override
    public String getColumn() {
        return columnName;
    }

    protected abstract Function<M, String> getFieldNameFunction();

    protected abstract Function<T,Object> getAccessor(M annotatedElement);

    void setColumnName(String columnName) {
        this.columnName = columnName;
    }
}
