package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;

import java.util.Map;
import java.util.SortedMap;
import java.util.stream.Collectors;

public interface ReportConfiguration<T> {

    SortedMap<String, ColumnConfiguration<T>> columns();

    String getReportSheetName();

    String getReportTitle();

    String getReportTitleStyleName();

    String getReportDescriptionDetail();

    String getReportDescriptionDetailStyleName();

    CellGenerator applyStyleAndValueToCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, Object transformedValue);

    void createStyles(CellStyler cellStyler);

    String getColumnHeaderStyleName();

    default ColumnConfiguration<T> getColumnConfiguration(String columnIdentifier) {
        return columns().get(columnIdentifier);
    }

    default Map<String, String> columnHeaders() {
        return columns().values()
            .stream()
            .collect(Collectors.toMap(ColumnConfiguration::getColumnName, ColumnConfiguration::getColumnHeader));
    }

    boolean hasReportDescriptionDetails();
}
