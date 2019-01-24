package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.TableStyle;
import com.wesleyhome.poi.api.internal.TableConfiguration;
import com.wesleyhome.poi.api.report.annotations.Report;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import com.wesleyhome.poi.api.report.annotations.TotalsRowFunction;

import java.lang.reflect.AnnotatedElement;
import java.util.List;
import java.util.stream.Collectors;

public class AnnotatedTableConfiguration implements TableConfiguration {


    private final TableStyle tableStyle;
    private final List<TotalsRowFunction> totalsRowFunctions;
    private final boolean hasTotalRow;

    public AnnotatedTableConfiguration(Report reportAnnotation, List<AnnotatedElement> totalsRowFunctions) {
        this.tableStyle = reportAnnotation.tableStyle();
        this.totalsRowFunctions = totalsRowFunctions.stream()
            .map(m -> m.getAnnotation(ReportColumn.class))
            .map(ReportColumn::totalFunction)
            .collect(Collectors.toList());
        this.hasTotalRow = this.totalsRowFunctions.stream().anyMatch(f -> !TotalsRowFunction.F_NONE.equals(f));
    }

    @Override
    public TableStyle getTableStyle() {
        return tableStyle;
    }

    @Override
    public List<TotalsRowFunction> getTotalRowFunctions() {
        return totalsRowFunctions;
    }

}
