package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.report.annotations.Report;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import org.apache.commons.lang3.reflect.ConstructorUtils;
import org.apache.poi.ss.usermodel.IndexedColors;
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

public class AnnotatedReportConfiguration<T> implements ReportConfiguration<T> {
    private final Class<T> reportClass;
    private final ReportStyler reportStyler;
    private final String sheetName;
    private final String reportDescription;
    private final SafeLazyInitializer<SortedMap<String, ColumnConfiguration<T>>> columnInitializer;

    public AnnotatedReportConfiguration(Class<T> reportClass) {
        this.reportClass = reportClass;
        Report annotation = this.reportClass.getAnnotation(Report.class);
        if (annotation == null) {
            throw new NullPointerException("Report annotation required");
        }
        String _sheetName = annotation.sheetName();
        String _description = annotation.description();
        sheetName = Report.NULL.equals(_sheetName) ? this.reportClass.getSimpleName() : _sheetName;
        reportDescription = Report.NULL.equals(_description) ? null : _description;
        try {
            reportStyler = ConstructorUtils.invokeConstructor(annotation.styler());
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
        this.columnInitializer = new SafeLazyInitializer<SortedMap<String, ColumnConfiguration<T>>>() {

            @Override
            protected SortedMap<String, ColumnConfiguration<T>> safeInitialize() {
                //noinspection unchecked
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
        };
    }

    @Override
    public SortedMap<String, ColumnConfiguration<T>> columns() {
        return columnInitializer.get();
    }

    @Override
    public String getReportSheetName() {
        return this.sheetName;
    }

    @Override
    public String getReportDescription() {
        return this.reportDescription == null ? getReportSheetName() : this.reportDescription;
    }

    @Override
    public IndexedColors headerBackgroundColor() {
        return reportStyler.headerBackgroundColor();
    }

    @Override
    public IndexedColors headerFontColor() {
        return reportStyler.headerFontColor();
    }

    @Override
    public boolean isHeaderBold() {
        return reportStyler.isHeaderBold();
    }

    @Override
    public IndexedColors evenRowBackgroundColor() {
        return reportStyler.evenRowBackgroundColor();
    }

    @Override
    public IndexedColors evenRowFontColor() {
        return reportStyler.evenRowFontColor();
    }

    @Override
    public boolean isEvenRowBold() {
        return reportStyler.isEvenRowBold();
    }

    @Override
    public IndexedColors oddRowBackgroundColor() {
        return reportStyler.oddRowBackgroundColor();
    }

    @Override
    public IndexedColors oddRowFontColor() {
        return reportStyler.oddRowFontColor();
    }

    @Override
    public boolean isOddRowBold() {
        return reportStyler.isOddRowBold();
    }
}
