package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.util.Set;
import java.util.TreeSet;

public class DefaultSheetGenerator implements SheetGenerator {
    private final WorkbookGenerator workbookGenerator;
    private String sheetName;
    private RowGenerator workingRow;
    private int startRowNum = 0;
    private final ExtendedMap<Integer, DefaultRowGenerator> rows;
    private Set<Integer> autosizeColumns;
    private Set<Integer> hiddenColumns;
    private Table table;
    private TableConfiguration tableConfiguration;

    public DefaultSheetGenerator(WorkbookGenerator workbookGenerator, String sheetName) {
        this.workbookGenerator = workbookGenerator;
        this.sheetName = sheetName;
        this.rows = new ExtendedTreeMap<>();
        this.autosizeColumns = new TreeSet<>();
        this.hiddenColumns = new TreeSet<>();
    }

    @Override
    public SheetGenerator withSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    @Override
    public SheetGenerator withStartRow(int rowNum) {
        if (!rows.isEmpty()) {
            throw new IllegalArgumentException("Rows have already been generated. This must be called before any rows are created");
        }
        startRowNum = rowNum;
        return this;
    }

    @Override
    public Workbook createWorkbook() {
        return workbookGenerator.createWorkbook();
    }

    @Override
    public WorkbookGenerator workbook() {
        return this.workbookGenerator;
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        return workbookGenerator.sheet(sheetName);
    }

    @Override
    public SheetGenerator sheet() {
        return this;
    }

    @Override
    public RowGenerator row(int rowNum) {
        return workingRow = rows.computeIfAbsent(rowNum, rn -> new DefaultRowGenerator(this, rn));
    }

    @Override
    public RowGenerator row() {
        if (this.workingRow == null) {
            return nextRow();
        }
        return this.workingRow;
    }

    @Override
    public RowGenerator nextRow() {
        return row(getNextRowNum());
    }

    private int getNextRowNum() {
        if (rows.isEmpty()) {
            return startRowNum;
        }
        return rows.lastKey() + 1;
    }

    @Override
    public CellGenerator nextCell() {
        if (workingRow == null) {
            return nextRow().nextCell();
        }
        return workingRow.nextCell();
    }

    @Override
    public CellGenerator cell(int rowNum, int columnNumber) {
        return row(rowNum).cell(columnNumber);
    }

    @Override
    public CellGenerator cell() {
        return this.row().cell();
    }

    @Override
    public CellStyler cellStyler() {
        return this.workbookGenerator.cellStyler();
    }

    @Override
    public SheetGenerator autosize(int columnNum) {
        this.autosizeColumns.add(columnNum);
        return this;
    }

    @Override
    public SheetGenerator hide(int columnNum) {
        this.hiddenColumns.add(columnNum);
        return this;
    }

    @Override
    public int rowNum() {
        return workingRow.rowNum();
    }

    @Override
    public int columnNum() {
        return workingRow.columnNum();
    }

    @Override
    public CellGenerator startTable() {
        CellGenerator currentCell = this.cell();
        switch (getWorkbookType()) {
            case EXCEL_OPEN:
                table = new Table(currentCell);
                break;
            default:
                System.err.printf("%s don't have the ability to createSheet tables.%n", getWorkbookType());
                break;
        }
        return currentCell;
    }

    @Override
    public CellGenerator endTable(TableConfiguration tableConfiguration) {
        this.tableConfiguration = tableConfiguration;
        CellGenerator currentCell = this.cell();
        switch (getWorkbookType()) {
            case EXCEL_OPEN:
                    table.setEndCell(currentCell);
                break;
            default:
                System.err.printf("%s don't have the ability to createSheet tables.%n", getWorkbookType());
                break;
        }
        return currentCell;
    }


    @Override
    public WorkbookType getWorkbookType() {
        return this.workbookGenerator.getWorkbookType();
    }


    public void applySheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet(this.sheetName);
        rows.values()
            .forEach(rowGen -> rowGen.applyRow(sheet));
        autosizeColumns.forEach(colNum -> sheet.autoSizeColumn(colNum, true));
        hiddenColumns.forEach(colNum -> sheet.setColumnHidden(colNum, true));
        if (this.table != null && this.table.isValid()) {
            switch (getWorkbookType()) {
                case EXCEL_OPEN:
                    XSSFSheet xssfSheet = (XSSFSheet) sheet;
                    createTable(xssfSheet);
                    break;
                case EXCEL_STREAM:
                default:
                    break;
            }
        }
    }

    private void createTable(XSSFSheet xssfSheet) {
        TableStyle tableStyle = this.tableConfiguration.getTableStyle();
        boolean hasTotalRow = this.tableConfiguration.hasTotalRow();
        XSSFTable table = xssfSheet.createTable(this.table.getAreaReference(false, xssfSheet.getWorkbook()));
        String tableStyleString = tableStyle.toString();
        CTTable ctTable = table.getCTTable();
        ctTable.addNewAutoFilter();
//        if(hasTotalRow) {
//            ctTable.setTotalsRowCount(1L);
//            ctTable.setTotalsRowShown(true);
//            CTTableColumns tableColumns = ctTable.getTableColumns();
//            List<CTTableColumn> tableColumnList = tableColumns.getTableColumnList();
//            ListIterator<CTTableColumn> itr = tableColumnList.listIterator();
//            List<TotalsRowFunction> totalRowFunctions = this.tableConfiguration.getTotalRowFunctions();
//            while(itr.hasNext()){
//                TotalsRowFunction function = totalRowFunctions.get(itr.nextIndex());
//                CTTableColumn column = itr.next();
//                if(!TotalsRowFunction.F_NONE.equals(function)) {
//                    column.setTotalsRowFunction(function.getFunctionEnum());
//                    column.setT
//                }
//            }
//        }
        table.setName(xssfSheet.getSheetName()+"_Table1");
        table.setDisplayName(xssfSheet.getSheetName()+"_Data_Table");
        CTTableStyleInfo tableStyleInfo = ctTable.addNewTableStyleInfo();
        tableStyleInfo.setName(tableStyleString);
        XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
        style.setShowColumnStripes(false);
        style.setShowRowStripes(true);
        style.setFirstColumn(false);
        style.setLastColumn(false);
        style.setShowRowStripes(true);
    }

    @Override
    public String toString() {
        return this.sheetName;
    }
}
