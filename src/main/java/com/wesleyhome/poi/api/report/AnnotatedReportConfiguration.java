package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.report.annotations.Report;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import org.apache.commons.lang3.reflect.ConstructorUtils;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import static java.util.stream.Collectors.collectingAndThen;
import static java.util.stream.Collectors.toMap;

public class AnnotatedReportConfiguration<T> extends AbstractReportConfiguration<T> {
    private final Class<T> reportClass;
    private final Report reportAnnotation;

    public AnnotatedReportConfiguration(Class<T> reportClass) {
        this.reportClass = reportClass;
        reportAnnotation = this.reportClass.getAnnotation(Report.class);
        if (reportAnnotation == null) {
            throw new NullPointerException("Report annotation required");
        }
    }

    @Override
    protected String initializeSheetName() {
        String _sheetName = reportAnnotation.sheetName();
        return Report.NULL.equals(_sheetName) ? this.reportClass.getSimpleName() : _sheetName;
    }

    @Override
    protected String initializeReportDescription() {
        String _description = reportAnnotation.description();
        return Report.NULL.equals(_description) ? getReportSheetName() : _description;
    }

    @Override
    protected ReportStyler initializeReportStyler() {
        try {
            return ConstructorUtils.invokeConstructor(reportAnnotation.styler());
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    protected SortedMap<String, ColumnConfiguration<T>> initializeColumns() {
        List<AbstractAnnotatedColumnConfiguration<T, ? extends AccessibleObject>> columnConfigurations = ReflectionHelper.getAnnotatedMembers(reportClass, ReportColumn.class)
            .stream()
            .map(member -> member instanceof Field ? new FieldColumnConfiguration<T>((Field) member) : new MethodColumnConfiguration<T>((Method) member))
            .sorted()
            .collect(Collectors.toList());
        AtomicInteger nextColumn = columnConfigurations.stream()
            .filter(c -> !ReportColumn.NULL.equals(c.getColumnName()))
            .max(Comparator.naturalOrder())
            .map(AbstractAnnotatedColumnConfiguration::getColumnName)
            .map(CellReference::convertColStringToIndex)
            .map(AtomicInteger::new)
            .orElseGet(() -> new AtomicInteger(CellReference.convertColStringToIndex("A") - 1));
        columnConfigurations.stream()
            .filter(c -> ReportColumn.NULL.equals(c.getColumnName()))
            .forEach(c -> {
                int nextColumnRow = nextColumn.incrementAndGet();
                String nextColumnName = CellReference.convertNumToColString(nextColumnRow);
                c.setColumnName(nextColumnName);
            });
        return columnConfigurations.stream()
            .map(c -> (ColumnConfiguration<T>) c)
            .collect(collectingAndThen(
                toMap(ColumnConfiguration::getColumnName, c -> c),
                TreeMap::new
            ));
    }
}
