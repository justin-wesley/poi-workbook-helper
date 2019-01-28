package com.wesleyhome.poi.api.report;

import com.google.common.base.CaseFormat;
import com.wesleyhome.poi.api.internal.Hyperlink;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.WordUtils;

import java.io.UnsupportedEncodingException;
import java.lang.reflect.AnnotatedElement;
import java.lang.reflect.Member;
import java.net.URLEncoder;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.commons.lang3.reflect.ConstructorUtils.invokeConstructor;

public abstract class AbstractAnnotatedColumnConfiguration<T, M extends Member & AnnotatedElement> implements ColumnConfiguration<T> {
    private static final String PROPERTY_NAME_PH = "\\$\\{([\\w|\\d|_]+)\\}";
    private static final Pattern PROPERTY_PATTERN = Pattern.compile(PROPERTY_NAME_PH);
    private final ReportColumn reportColumn;
    private final M annotatedElement;
    private final ReportConfiguration<T> reportConfiguration;
    private final boolean display;
    private String columnName;
    private final SafeLazyInitializer<Function<T, Object>> accessorInitializer;
    private final SafeLazyInitializer<String> propertyNameInitializer;
    private final SafeLazyInitializer<String> defaultHeaderInitializer;
    private SafeLazyInitializer<Object> columnValueInitializer;

    public AbstractAnnotatedColumnConfiguration(M annotatedElement, ReportConfiguration<T> reportConfiguration) {
        this.reportColumn = annotatedElement.getAnnotation(ReportColumn.class);
        this.display = this.reportColumn.display();
        this.annotatedElement = annotatedElement;
        this.reportConfiguration = reportConfiguration;
        this.columnName = reportColumn.columnIdentifier();

        this.accessorInitializer = initializer(m->getAccessor(annotatedElement));
        this.propertyNameInitializer = initializer(m-> getPropertyNameFunction().apply(annotatedElement));
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
    public boolean isDisplayed() {
        return this.display;
    }

    @Override
    public String getPropertyName() {
        return this.propertyNameInitializer.get();
    }

    @Override
    public Object getColumnValue(T value) {
        Object columnValue = this.accessorInitializer.get().apply(value);
        return ColumnType.URL.equals(reportColumn.columnType()) ? createHyperlink(reportColumn, value, columnValue) : columnValue;
    }

    private Hyperlink createHyperlink(ReportColumn reportColumn, T value, Object columnValue) {
        String urlFormattingText = reportColumn.urlFormattingText();
        String columnValueString = columnValue.toString();
        if(isNull(urlFormattingText)){
            return new Hyperlink(columnValueString, columnValueString);
        }
        Matcher matcher = PROPERTY_PATTERN.matcher(urlFormattingText);
        while(matcher.find()){

            String propertyName = matcher.group(1);
            if("this".equals(propertyName)){
                urlFormattingText = urlFormattingText.replace("${"+propertyName+"}", columnValueString);
            }else {
                ColumnConfiguration<T> columnConfiguration = reportConfiguration.getColumnConfiguration(propertyName);
                Object propertyColumnValue = columnConfiguration.getColumnValue(value);
                urlFormattingText = urlFormattingText.replace("${"+propertyName+"}", propertyColumnValue.toString());
            }
        }
            return new Hyperlink(urlFormattingText, columnValueString);
    }

    protected boolean isNull(String value) {
        return value == null || value.trim().length() == 0 || ReportColumn.NULL.equals(value);
    }

    @Override
    public ColumnType getColumnType() {
        return this.reportColumn.columnType();
    }

    @Override
    public String getColumnHeader() {
        return ReportColumn.NULL.equals(reportColumn.columnHeader()) ? StringUtils.capitalize(getDefaultHeaderName()) : reportColumn.columnHeader();
    }

    protected String getDefaultHeaderName(){
        return defaultHeaderInitializer.get();
    }

    protected Function<M, String> getDefaultColumnHeaderNameFunction() {
        return field-> WordUtils.capitalize(CaseFormat.LOWER_CAMEL.to(
            CaseFormat.UPPER_UNDERSCORE,
            propertyNameInitializer.get()
        ).replace("_", " "));
    }

    @Override
    public String getColumn() {
        return columnName;
    }

    protected abstract Function<M, String> getPropertyNameFunction();

    protected abstract Function<T,Object> getAccessor(M annotatedElement);

    void setColumnName(String columnName) {
        this.columnName = columnName;
    }
}
