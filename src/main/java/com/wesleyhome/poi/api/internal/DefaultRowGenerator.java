package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultRowGenerator implements RowGenerator {
    private final SheetGenerator sheetGenerator;
    private final int rowNum;
    private CellGenerator currentCell;
    private ExtendedMap<Integer, DefaultCellGenerator> cells;
    private int startColumn;

    public DefaultRowGenerator(SheetGenerator sheetGenerator, int rowNum) {
        this.sheetGenerator = sheetGenerator;
        this.rowNum = rowNum;
        this.cells = new ExtendedTreeMap<>();
        this.startColumn = 0;
    }

    @Override
    public RowGenerator withStartColumn(int columnNum) {
        if(!cells.isEmpty()){
            throw new IllegalArgumentException("Must call this before creating cells");
        }
        this.startColumn = columnNum;
        return this;
    }

    @Override
    public Workbook createWorkbook() {
        return sheetGenerator.createWorkbook();
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        return sheetGenerator.sheet(sheetName);
    }

    @Override
    public RowGenerator nextRow() {
        return sheetGenerator.nextRow();
    }

    @Override
    public RowGenerator row(int rowNum) {
        return sheetGenerator.row(rowNum);
    }

    @Override
    public CellGenerator nextCell() {
        return cell(getNextColumnNum());
    }

    private int getNextColumnNum() {
        if(cells.isEmpty()){
            return startColumn;
        }
        return cells.lastKey()+1;
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNum) {
        return this.sheetGenerator.cell(columnNumber, rowNum);
    }

    @Override
    public CellGenerator cell(int columnNum) {
        return currentCell = cells.getOrDefault(columnNum, () -> new DefaultCellGenerator(this, columnNum));
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.sheetGenerator.cellStyleManager();
    }

    public void applyRow(Sheet sheet) {
        Row row = sheet.createRow(this.rowNum);
        cells.values()
            .forEach(cellGen -> {
                cellGen.applyCell(row);
            });
    }
}
