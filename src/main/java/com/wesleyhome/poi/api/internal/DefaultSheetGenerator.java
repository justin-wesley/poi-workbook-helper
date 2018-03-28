package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.atomic.AtomicInteger;

public class DefaultSheetGenerator implements SheetGenerator {
    private final WorkbookGenerator workbookGenerator;
    private String sheetName;
    private RowGenerator workingRow;
    private final AtomicInteger currentRowNum;
    private int startRowNum = 0;
    private final Map<Integer, RowGenerator> rows;

    public DefaultSheetGenerator(WorkbookGenerator workbookGenerator, String sheetName) {
        this.workbookGenerator = workbookGenerator;
        this.sheetName = sheetName;
        this.currentRowNum = new AtomicInteger();
        this.rows = new TreeMap<>();
    }

    @Override
    public SheetGenerator withSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    @Override
    public SheetGenerator withStartRow(int rowNum) {
        if(!rows.isEmpty()){
            throw new IllegalArgumentException("Rows have already been generated. This must be called before any rows are created");
        }
        startRowNum = rowNum;
        currentRowNum.set(startRowNum);
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
        workingRow = rows.getOrDefault(rowNum, new DefaultRowGenerator(this, rowNum));
        currentRowNum.set(rowNum);
        return workingRow;
    }

    @Override
    public RowGenerator nextRow() {
        return row(currentRowNum.incrementAndGet());
    }

    @Override
    public CellGenerator nextCell() {
        return workingRow.nextCell();
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNum) {
        return row(rowNum).cell(columnNumber);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.workbookGenerator.cellStyleManager();
    }
}
