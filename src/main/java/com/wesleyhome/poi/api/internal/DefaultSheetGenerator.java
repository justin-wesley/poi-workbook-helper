package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultSheetGenerator implements SheetGenerator {
    private final WorkbookGenerator workbookGenerator;
    private String sheetName;
    private RowGenerator workingRow;
    private int startRowNum = 0;
    private final ExtendedMap<Integer, DefaultRowGenerator> rows;

    public DefaultSheetGenerator(WorkbookGenerator workbookGenerator, String sheetName) {
        this.workbookGenerator = workbookGenerator;
        this.sheetName = sheetName;
        this.rows = new ExtendedTreeMap<>();
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
    public SheetGenerator sheet(String sheetName) {
        return workbookGenerator.sheet(sheetName);
    }

    @Override
    public RowGenerator row(int rowNum) {
        return workingRow = rows.computeIfAbsent(rowNum, rn -> new DefaultRowGenerator(this, rn));
    }

    @Override
    public RowGenerator row() {
        if(this.workingRow == null){
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
    public CellStyleManager cellStyleManager() {
        return this.workbookGenerator.cellStyleManager();
    }

    public void applySheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet(this.sheetName);
        rows.values()
            .forEach(rowGen -> rowGen.applyRow(sheet));
    }

    @Override
    public int rowNum() {
        return workingRow.rowNum();
    }

    @Override
    public int columnNum() {
        return workingRow.columnNum();
    }
}
