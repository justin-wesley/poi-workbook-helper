package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

public class DefaultReportGenerator<T> implements ReportGenerator<T> {

    @Override
    public Workbook generateWorkbook(Iterable<T> data, ReportConfiguration<T> reportConfiguration) {
        Map<String, String> columnHeaderMap = reportConfiguration.columnHeaders();
        return WorkbookGenerator.create(reportConfiguration.getReportSheetName())
            .generateStyles(reportConfiguration::createStyles)
            .nextCell()
            .usingStyle(reportConfiguration.getReportTitleStyleName())
            .mergeWithNextXCells(reportConfiguration.columns().size()-1)
            .havingValue(reportConfiguration.getReportTitle())
            .nextRow().nextCell()
            .applyIf(reportConfiguration.hasReportDescriptionDetails(), cg->
                cg.usingStyle(reportConfiguration.getReportDescriptionDetailStyleName())
                    .mergeWithNextXCells(reportConfiguration.columns().size()-1)
                    .havingValue(reportConfiguration.getReportDescriptionDetail())
                .nextRow().cell()
            )
            .nextRow()
            .startTable()
            .generateCells(columnHeaderMap.values(), ((cellGenerator, headerName) -> cellGenerator.autosize().havingValue(headerName))).nextRow()
            .generateRows(data, ((rowGenerator, value) -> rowGenerator
                .generateCells(columnHeaderMap.keySet(), ((cellGenerator, columnIdentifier) -> this.generateValueCell(cellGenerator, columnIdentifier, reportConfiguration, value)))
                .row()))
            .endTable(reportConfiguration.getTableConfiguration())
            .createWorkbook();
    }

    private CellGenerator generateValueCell(CellGenerator cellGenerator, String columnIdentifier, ReportConfiguration<T> reportConfiguration, T value) {
        ColumnConfiguration<T> columnConfiguration = reportConfiguration.getColumnConfiguration(columnIdentifier);
        Object transformedValue = columnConfiguration.getColumnValue(value);
        return reportConfiguration.applyStyleAndValueToCell(cellGenerator, columnConfiguration, transformedValue);
    }
}
