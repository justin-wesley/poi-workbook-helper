package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFFormulaEvaluator;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DefaultWorkbookGenerator implements WorkbookGenerator {

    private final CellStyleManager cellStyleManager;
    private final WorkbookType workbookType;
    private ExtendedMap<String, DefaultSheetGenerator> sheets;
    private DefaultSheetGenerator workingSheet;

    public DefaultWorkbookGenerator(WorkbookType workbookType) {
        this.workbookType = workbookType;
        sheets = new ExtendedTreeMap<>();
        this.cellStyleManager = new CellStyleManager();
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
        return workingSheet = sheets.computeIfAbsent(safeSheetName, name->new DefaultSheetGenerator(this, name));
    }

    @Override
    public Workbook createWorkbook() {
        Workbook workbook = createNewWorkbook();
        sheets.values()
            .forEach(sheetGen -> sheetGen.applySheet(workbook));
//        FormulaEvaluator evaluator = getEvaluator(workbook);
//        evaluator.evaluateAll();
        return workbook;
    }

    private FormulaEvaluator getEvaluator(Workbook workbook) {
        switch (workbookType){
            case EXCEL_BIN:
                return new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
            case EXCEL_OPEN:
            default:
                return new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
            case EXCEL_STREAM:
                return new SXSSFFormulaEvaluator((SXSSFWorkbook) workbook);
        }
    }


//    private void evaluateFormula(Workbook workbook) {
//    }

    private Workbook createNewWorkbook() {
        switch (this.workbookType) {
            case EXCEL_BIN:
                return new HSSFWorkbook();
            case EXCEL_STREAM:
                return new SXSSFWorkbook();
            default:
                return new XSSFWorkbook();
        }
    }

    @Override
    public RowGenerator nextRow() {
        return workingSheet.nextRow();
    }

    @Override
    public RowGenerator row() {
        return this.workingSheet.row();
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
    public CellGenerator cell() {
        return workingSheet.cell();
    }

    @Override
    public CellGenerator cell(int rowNumber, int columnNumber) {
        return workingSheet.cell(rowNumber, columnNumber);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return cellStyleManager;
    }

    @Override
    public int rowNum() {
        return workingSheet.rowNum();
    }

    @Override
    public int columnNum() {
        return workingSheet.columnNum();
    }
}
