package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class DefaultCellGenerator implements CellGenerator {
    private final RowGenerator rowGenerator;
    private final int columnNum;
    private CellStyle cellStyle;
    private Object cellValue;

    public DefaultCellGenerator(RowGenerator rowGenerator, int columnNum) {
        this.rowGenerator = rowGenerator;
        this.columnNum = columnNum;
        this.cellStyle = new CellStyle();
    }

    @Override
    public CellGenerator havingValue(Object cellValue) {
        this.cellValue = cellValue;
        return this;
    }

    @Override
    public CellGenerator usingStyle(String cellStyle) {
        this.cellStyle = cellStyleManager().copy(cellStyle);
        return this;
    }

    @Override
    public CellGenerator withDateFormat() {
        return this;
    }

    @Override
    public CellGenerator withIntegerFormat() {
        return this;
    }

    @Override
    public CellGenerator withCurrencyFormat() {
        return this;
    }

    @Override
    public CellGenerator withNumericStyle() {
        return this;
    }

    @Override
    public CellGenerator withBackgroundColor(IndexedColors color) {
        return this;
    }

    @Override
    public CellGenerator withFontColor(IndexedColors color) {
        return this;
    }

    @Override
    public CellGenerator isBold() {
        return this;
    }

    @Override
    public CellGenerator withNoBackgroundColor() {
        return this;
    }

    @Override
    public CellGenerator as(String cellStyleName) {
        return this;
    }

    @Override
    public Workbook createWorkbook() {
        return this.rowGenerator.createWorkbook();
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        return this.rowGenerator.sheet(sheetName);
    }

    @Override
    public RowGenerator nextRow() {
        return this.rowGenerator.nextRow();
    }

    @Override
    public RowGenerator row(int rowNum) {
        return this.rowGenerator.row(rowNum);
    }

    @Override
    public CellGenerator nextCell() {
        return this.rowGenerator.nextCell();
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNum) {
        return this.rowGenerator.cell(columnNumber, rowNum);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.rowGenerator.cellStyleManager();
    }
}
