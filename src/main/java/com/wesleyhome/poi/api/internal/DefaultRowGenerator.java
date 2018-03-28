package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.TreeMap;

public class DefaultRowGenerator implements RowGenerator {
    private final SheetGenerator sheetGenerator;
    private final int rowNum;
    private int startColumn;
    private CellGenerator currentCell;
    private int currentColumnNum;
    private Map<Integer, CellGenerator> cells;

    public DefaultRowGenerator(SheetGenerator sheetGenerator, int rowNum) {
        this.sheetGenerator = sheetGenerator;
        this.rowNum = rowNum;
        cells = new TreeMap<>();
    }

    @Override
    public RowGenerator withStartColumn(int columnNum) {
        if(!cells.isEmpty()){
            throw new IllegalArgumentException("Must call this before creating cells");
        }
        startColumn = columnNum;
        currentColumnNum = startColumn;
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
        return null;
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNum) {
        return this.sheetGenerator.cell(columnNumber, rowNum);
    }

    @Override
    public CellGenerator cell(int columnNum) {
        return cells.getOrDefault(columnNum, new DefaultCellGenerator(this, columnNum));
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.sheetGenerator.cellStyleManager();
    }
}
