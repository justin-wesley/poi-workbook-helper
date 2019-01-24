package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;
import com.wesleyhome.poi.api.internal.TableConfiguration;

import java.util.List;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.stream.Collectors;

public interface ReportConfiguration<T> {

    String getReportSheetName();

    String getReportTitle();

    String getReportTitleStyleName();

    String getReportDescriptionDetail();

    String getReportDescriptionDetailStyleName();

    CellGenerator applyStyleAndValueToCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, Object transformedValue);

    void createStyles(CellStyler cellStyler);

    SortedSet<ColumnConfiguration<T>> columns();

    default List<String> columnHeaders() {
        return columns()
            .stream()
            .sorted()
            .map(ColumnConfiguration::getColumnHeader)
            .collect(Collectors.toList());
    }

    boolean hasReportDescriptionDetails();

    TableConfiguration getTableConfiguration();
}
