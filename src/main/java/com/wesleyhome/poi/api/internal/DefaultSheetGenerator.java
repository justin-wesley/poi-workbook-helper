package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.concurrent.atomic.AtomicInteger;

public class DefaultSheetGenerator implements SheetGenerator {
    private final WorkbookGenerator workbookGenerator;
    private String sheetName;
    private RowGenerator workingRow;
    private int startRowNum = 0;
    private final ExtendedMap<Integer, DefaultRowGenerator> rows;
    private Set<Integer> autosizeColumns;
    private Set<Integer> hiddenColumns;

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
        return this.workbookGenerator.startTable();
    }

    @Override
    public CellGenerator endTable(TableConfiguration tableConfiguration) {
        return this.workbookGenerator.endTable(tableConfiguration);
    }


    @Override
    public WorkbookType getWorkbookType() {
        return this.workbookGenerator.getWorkbookType();
    }


    public void applySheet(Workbook workbook, Map<Table, TableConfiguration> tables, AtomicInteger tableCount) {
        Sheet sheet = workbook.createSheet(this.sheetName);
        rows.values()
            .forEach(rowGen -> rowGen.applyRow(sheet));
        autosizeColumns.forEach(colNum -> sheet.autoSizeColumn(colNum, true));
        hiddenColumns.forEach(colNum -> sheet.setColumnHidden(colNum, true));
        if (!tables.isEmpty()) {
            switch (getWorkbookType()) {
                case EXCEL_OPEN:
                    XSSFSheet xssfSheet = (XSSFSheet) sheet;
                    createTable(xssfSheet, tables, tableCount);
                    break;
                case EXCEL_STREAM:
                default:
                    break;
            }
        }
    }

    private void createTable(XSSFSheet xssfSheet, Map<Table, TableConfiguration> tables, AtomicInteger tableCount) {
        tables.forEach(((table1, tableConfiguration) -> {
            TableStyle tableStyle = tableConfiguration.getTableStyle();
            boolean hasTotalRow = tableConfiguration.hasTotalRow();
            XSSFTable table = xssfSheet.createTable(table1.getAreaReference(false, xssfSheet.getWorkbook()));
            String tableStyleString = tableStyle.toString();
            CTTable ctTable = table.getCTTable();
            ctTable.addNewAutoFilter();
//            if (hasTotalRow) {
//                ctTable.setTotalsRowCount(1L);
//                ctTable.setTotalsRowShown(true);
//                CTTableColumns tableColumns = ctTable.getTableColumns();
//                List<CTTableColumn> tableColumnList = tableColumns.getTableColumnList();
//                ListIterator<CTTableColumn> itr = tableColumnList.listIterator();
//                List<TotalsRowFunction> totalRowFunctions = this.tableConfiguration.getTotalRowFunctions();
//                while (itr.hasNext()) {
//                    TotalsRowFunction function = totalRowFunctions.get(itr.nextIndex());
//                    CTTableColumn column = itr.next();
//                    if (!TotalsRowFunction.F_NONE.equals(function)) {
//                        column.setTotalsRowFunction(function.getFunctionEnum());
//                        column.setT
//                    }
//                }
//            }
            table.setName("Table" + tableCount.getAndIncrement());
            table.setDisplayName(table.getName());
            CTTableStyleInfo tableStyleInfo = ctTable.addNewTableStyleInfo();
            tableStyleInfo.setName(tableStyleString);
            XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
            style.setShowColumnStripes(false);
            style.setShowRowStripes(true);
            style.setFirstColumn(false);
            style.setLastColumn(false);
            style.setShowRowStripes(true);
        }));

    }

    @Override
    public String toString() {
        return this.sheetName;
    }
}
