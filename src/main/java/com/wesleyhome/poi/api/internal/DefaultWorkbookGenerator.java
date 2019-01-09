package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFFormulaEvaluator;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.stream.Collectors;
import java.util.stream.Stream;

public class DefaultWorkbookGenerator implements WorkbookGenerator {

    public interface BuiltinStyles {
        String ACCOUNTING = "ACCOUNTING";
        String CURRENCY = "CURRENCY";
        String DATE_TIME = "DATE_TIME";
        String INTEGER = "INTEGER";
        String NO_BACKGROUND_COLOR = "NO_BACKGROUND_COLOR";
        String NUMERIC = "NUMERIC";
        String WRAPPED_TEXT = "WRAPPED_TEXT";
        String BOLD = "BOLD";
        String DATE = "DATE";

        static String alignName(HorizontalAlignment ha) {
            return ha.name()+"_H_ALIGN";
        }
    }

    private final WorkbookType workbookType;
    private ExtendedMap<String, DefaultSheetGenerator> sheets;
    private DefaultSheetGenerator workingSheet;
    private CellStyler cellStyler;

    public DefaultWorkbookGenerator(WorkbookType workbookType) {
        this.workbookType = workbookType;
        sheets = new ExtendedTreeMap<>();
        this.cellStyler = new DefaultCellStyler().withAccountingFormat()
            .as(BuiltinStyles.ACCOUNTING)
            .reset()
            .withCurrencyFormat()
            .as(BuiltinStyles.CURRENCY)
            .reset()
            .withDateTimeFormat()
            .as(BuiltinStyles.DATE_TIME)
            .reset()
            .withDateFormat()
            .as((BuiltinStyles.DATE))
            .reset()
            .withIntegerFormat()
            .as(BuiltinStyles.INTEGER)
            .reset()
            .withNoBackgroundColor()
            .as(BuiltinStyles.NO_BACKGROUND_COLOR)
            .reset()
            .withNumericStyle()
            .as(BuiltinStyles.NUMERIC)
            .reset()
            .withWrappedText()
            .as(BuiltinStyles.WRAPPED_TEXT)
            .reset()
            .isBold()
            .as(BuiltinStyles.BOLD)
            .reset()
            .applyEach(HorizontalAlignment.values(), (cs, ha)->cs.withHorizontalAlignment(ha).as(BuiltinStyles.alignName(ha)));
    }

    @Override
    public CellStyler cellStyler() {
        return this.cellStyler;
    }

    @Override
    public WorkbookGenerator workbook() {
        return this;
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
    public int rowNum() {
        return workingSheet.rowNum();
    }

    @Override
    public int columnNum() {
        return workingSheet.columnNum();
    }

    @Override
    public String toString() {
        return this.sheets.values().stream().map(Object::toString).collect(Collectors.joining(","));
    }
}
