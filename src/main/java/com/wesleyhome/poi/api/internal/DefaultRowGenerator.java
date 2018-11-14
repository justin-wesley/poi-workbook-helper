package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.stream.Collectors;

public class DefaultRowGenerator implements RowGenerator {
    private final SheetGenerator sheetGenerator;
    private final int rowNum;
    private Float rowHeight;
    private DefaultCellGenerator workingCell;
    private NavigableMap<Integer, DefaultCellGenerator> cells;
    private int startColumn;

    public DefaultRowGenerator(SheetGenerator sheetGenerator, int rowNum) {
        this.sheetGenerator = sheetGenerator;
        this.rowNum = rowNum;
        this.cells = new TreeMap<>();
        this.startColumn = 0;
    }

    @Override
    public RowGenerator withStartColumn(int columnNum) {
        if (!cells.isEmpty()) {
            throw new IllegalArgumentException("Must call this before creating cells");
        }
        this.startColumn = columnNum;
        return this;
    }

    @Override
    public int rowNum() {
        return this.rowNum;
    }

    @Override
    public int columnNum() {
        return this.workingCell.columnNum();
    }

    @Override
    public Workbook createWorkbook() {
        return sheetGenerator.createWorkbook();
    }

    @Override
    public WorkbookGenerator workbook() {
        return this.sheetGenerator.workbook();
    }

    @Override
    public SheetGenerator sheet() {
        return sheetGenerator;
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
    public RowGenerator row() {
        return this;
    }

    @Override
    public CellGenerator nextCell() {
        return cell(getNextColumnNum());
    }

    @Override
    public CellGenerator cell() {
        return this.workingCell == null ? nextCell() : this.workingCell;
    }

    private int getNextColumnNum() {
        if (cells.isEmpty()) {
            return startColumn;
        }
        DefaultCellGenerator last = cells.lastEntry().getValue();
        int columnNum = last.columnNum();
        int cellsToMerge = last.getCellsToMerge();
        return columnNum + cellsToMerge + 1;
    }

    @Override
    public CellGenerator cell(int rowNum, int columnNumber) {
        return this.sheetGenerator.cell(rowNum, columnNumber);
    }

    @Override
    public CellGenerator cell(int columnNum) {
        return cells.computeIfAbsent(columnNum, colNum-> new DefaultCellGenerator(this, colNum));
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.sheetGenerator.cellStyleManager();
    }

    void applyRow(Sheet sheet) {
        Row row = sheet.createRow(this.rowNum);
        if (this.rowHeight != null) {
            row.setHeightInPoints(this.rowHeight);
        }
        cells.forEach((colNum, cellGen) -> cellGen.applyCell(row));
    }

    @Override
    public RowGenerator height(float rowHeight) {
        this.rowHeight = rowHeight;
        return this;
    }

    @Override
    public String toString() {
        return this.cells.values().stream().map(Object::toString).collect(Collectors.joining(","));
    }
}
