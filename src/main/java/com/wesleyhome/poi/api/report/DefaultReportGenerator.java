package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.*;

public class DefaultReportGenerator implements ReportGenerator {

    private final WorkbookGenerator workbookGenerator;
    private String currentSheetName;

    public DefaultReportGenerator() {
        workbookGenerator = WorkbookGenerator.createNewWorkbook();
    }

    @Override
    public <T> ReportGenerator applyReport(Iterable<T> data, ReportConfiguration<T> reportConfiguration) {
        if(currentSheetName != null && currentSheetName.equals(reportConfiguration.getReportSheetName())){
            workbookGenerator.sheet(currentSheetName).nextRow().nextRow();
        }
        generateSheet(data, reportConfiguration);
        currentSheetName = reportConfiguration.getReportSheetName();
        return this;
    }

    @Override
    public Workbook create() {
        return workbookGenerator.createWorkbook();
    }

    private <T> SheetGenerator generateSheet(Iterable<T> data, ReportConfiguration<T> reportConfiguration) {
        List<String> columnHeaders = reportConfiguration.columnHeaders();
        SortedSet<ColumnConfiguration<T>> displayedColumns = reportConfiguration.displayedColumns();
        return workbookGenerator.generateStyles(reportConfiguration::createStyles)
            .sheet(reportConfiguration.getReportSheetName())
            .nextCell()
            .applyIf(reportConfiguration.getReportTitle() != null, cg-> cg.usingStyle(reportConfiguration.getReportTitleStyleName())
                .mergeWithNextXCells(displayedColumns.size()-1)
                .havingValue(reportConfiguration.getReportTitle())
                .nextRow().nextCell())
            .applyIf(reportConfiguration.hasReportDescriptionDetails(), cg->
                cg.usingStyle(reportConfiguration.getReportDescriptionDetailStyleName())
                    .mergeWithNextXCells(displayedColumns.size()-1)
                    .havingValue(reportConfiguration.getReportDescriptionDetail())
                    .nextRow().nextCell()
            )
            .applyIf(reportConfiguration.getReportTitle() != null || reportConfiguration.hasReportDescriptionDetails(), cg->cg.nextRow().nextCell())
            .startTable()
            .generateCells(columnHeaders, ((cellGenerator, headerName) -> cellGenerator.autosize().havingValue(headerName))).nextRow()
            .generateRows(data, ((rowGenerator, value) -> rowGenerator
                .generateCells(displayedColumns, ((cellGenerator, columnConfiguration) -> this.generateValueCell(cellGenerator, columnConfiguration, reportConfiguration, value)))
                .row()))
            .endTable(reportConfiguration.getTableConfiguration())
            .sheet();

    }

    private <T> CellGenerator generateValueCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, ReportConfiguration<T> reportConfiguration, T value) {
        Object transformedValue = columnConfiguration.getColumnValue(value);
        return reportConfiguration.applyStyleAndValueToCell(cellGenerator, columnConfiguration, transformedValue);
    }
}
