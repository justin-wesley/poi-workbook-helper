package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.*;

public class DefaultReportGenerator implements ReportGenerator {

    private final WorkbookGenerator workbookGenerator;

    public DefaultReportGenerator() {
        workbookGenerator = WorkbookGenerator.createNewWorkbook();
    }

    @Override
    public <T> ReportGenerator applyReport(Iterable<T> data, ReportConfiguration<T> reportConfiguration) {
        generateSheet(data, reportConfiguration);
        return this;
    }

    @Override
    public Workbook create() {
        return workbookGenerator.createWorkbook();
    }

    private <T> SheetGenerator generateSheet(Iterable<T> data, ReportConfiguration<T> reportConfiguration) {
        Map<String, String> columnHeaderMap = reportConfiguration.columnHeaders();
        return workbookGenerator.generateStyles(reportConfiguration::createStyles)
            .sheet(reportConfiguration.getReportSheetName())
            .nextCell()
            .applyIf(reportConfiguration.getReportTitle() != null, cg-> cg.usingStyle(reportConfiguration.getReportTitleStyleName())
                .mergeWithNextXCells(reportConfiguration.columns().size()-1)
                .havingValue(reportConfiguration.getReportTitle())
                .nextRow().nextCell())
            .applyIf(reportConfiguration.hasReportDescriptionDetails(), cg->
                cg.usingStyle(reportConfiguration.getReportDescriptionDetailStyleName())
                    .mergeWithNextXCells(reportConfiguration.columns().size()-1)
                    .havingValue(reportConfiguration.getReportDescriptionDetail())
                    .nextRow().nextCell()
            )
            .applyIf(reportConfiguration.getReportTitle() != null || reportConfiguration.hasReportDescriptionDetails(), cg->cg.nextRow().nextCell())
            .startTable()
            .generateCells(columnHeaderMap.values(), ((cellGenerator, headerName) -> cellGenerator.autosize().havingValue(headerName))).nextRow()
            .generateRows(data, ((rowGenerator, value) -> rowGenerator
                .generateCells(columnHeaderMap.keySet(), ((cellGenerator, columnIdentifier) -> this.generateValueCell(cellGenerator, columnIdentifier, reportConfiguration, value)))
                .row()))
            .endTable(reportConfiguration.getTableConfiguration())
            .sheet();

    }

    private <T> CellGenerator generateValueCell(CellGenerator cellGenerator, String columnIdentifier, ReportConfiguration<T> reportConfiguration, T value) {
        ColumnConfiguration<T> columnConfiguration = reportConfiguration.getColumnConfiguration(columnIdentifier);
        Object transformedValue = columnConfiguration.getColumnValue(value);
        return reportConfiguration.applyStyleAndValueToCell(cellGenerator, columnConfiguration, transformedValue);
    }
}
