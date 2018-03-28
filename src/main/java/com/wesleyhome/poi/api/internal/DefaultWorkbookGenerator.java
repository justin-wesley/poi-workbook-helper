package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.util.Map;
import java.util.TreeMap;

public class DefaultWorkbookGenerator implements WorkbookGenerator {

    private final CellStyleManager cellStyleManager;
    private Map<String, SheetGenerator> sheets;
    private SheetGenerator workingSheet;

    public DefaultWorkbookGenerator(){
        sheets = new TreeMap<>();
        this.cellStyleManager = new CellStyleManager();
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
        workingSheet = sheets.getOrDefault(safeSheetName, new DefaultSheetGenerator(this, safeSheetName));
        return workingSheet;
    }

    @Override
    public Workbook createWorkbook() {
        return null;
    }

    @Override
    public RowGenerator nextRow() {
        return workingSheet.nextRow();
    }

    @Override
    public RowGenerator row(int rowNum) {
        return workingSheet.row(rowNum);
    }

    @Override
    public CellGenerator nextCell() {
        return workingSheet.nextCell();
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNumber) {
        return workingSheet.cell(columnNumber, rowNumber);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return cellStyleManager;
    }
}
