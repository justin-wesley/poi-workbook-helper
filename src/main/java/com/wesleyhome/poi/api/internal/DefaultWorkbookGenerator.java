package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFFormulaEvaluator;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;
import java.util.stream.Collectors;

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
        String ZIP = "ZIP";
        String URL = "URL";

        static String alignName(HorizontalAlignment ha) {
            return ha.name()+"_H_ALIGN";
        }
    }

    private final WorkbookType workbookType;
    private Map<String, DefaultSheetGenerator> sheets;
    private DefaultSheetGenerator workingSheet;
    private CellStyler cellStyler;
    private Table currentTable;
    private Map<DefaultSheetGenerator, Map<Table, TableConfiguration>> tables;

    public DefaultWorkbookGenerator(WorkbookType workbookType) {
        this.workbookType = workbookType;
        sheets = new LinkedHashMap<>();
        tables = new HashMap<>();
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
            .withNumericFormat()
            .as(BuiltinStyles.NUMERIC)
            .reset()
            .withWrappedText()
            .as(BuiltinStyles.WRAPPED_TEXT)
            .reset()
            .withBoldFont()
            .as(BuiltinStyles.BOLD)
            .reset()
            .withZipCodeFormat()
            .as(BuiltinStyles.ZIP)
            .reset()
            .withUnderline()
            .withFontColor(IndexedColors.BLUE)
            .as(BuiltinStyles.URL)
            .reset()
            .applyEach(HorizontalAlignment.values(), (cs, ha)->cs.withHorizontalAlignment(ha).as(BuiltinStyles.alignName(ha)));
    }

    @Override
    public WorkbookGenerator generateStyles(Consumer<CellStyler> cellStyler) {
        cellStyler.accept(cellStyler());
        return this;
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
    public SheetGenerator sheet() {
        if(workingSheet == null) {
            return sheet("Sheet1");
        }
        return workingSheet;
    }

    @Override
    public Workbook createWorkbook() {
        Workbook workbook = createNewWorkbook();

        AtomicInteger tableCount = new AtomicInteger(1);
        sheets.values()
            .forEach(sheetGen -> sheetGen.applySheet(workbook, tables.computeIfAbsent(sheetGen, sg->new HashMap<>()), tableCount));
        return workbook;
    }

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
    public WorkbookType getWorkbookType() {
        return workbookType;
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

    @Override
    public CellGenerator startTable() {
        CellGenerator currentCell = this.cell();
        if (currentTable != null) {
            System.err.printf("You must end a table before creating a new one");
        } else {
            switch (getWorkbookType()) {
                case EXCEL_OPEN:
                    currentTable = new Table(currentCell);
                    break;
                default:
                    System.err.printf("%s don't have the ability to createSheet tables.%n", getWorkbookType());
                    break;
            }
        }
        return currentCell;
    }

    @Override
    public CellGenerator endTable(TableConfiguration tableConfiguration) {
        CellGenerator currentCell = this.cell();
        switch (getWorkbookType()) {
            case EXCEL_OPEN:
                currentTable.setEndCell(currentCell);
                Map<Table, TableConfiguration> tableMap = tables.computeIfAbsent((DefaultSheetGenerator) currentCell.sheet(), (s) -> new HashMap<>());
                tableMap.put(currentTable, tableConfiguration);
                currentTable = null;
                break;
            default:
                System.err.printf("%s don't have the ability to createSheet tables.%n", getWorkbookType());
                break;
        }
        return currentCell;
    }
}
