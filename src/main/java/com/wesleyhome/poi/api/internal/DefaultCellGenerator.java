package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import org.apache.poi.ss.usermodel.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class DefaultCellGenerator implements CellGenerator {
    private final RowGenerator rowGenerator;
    private final int columnNum;
    private ExtendedCellStyle extendedCellStyle;
    private Object cellValue;
    private CellType cellType;

    enum CellType {
        DATE,
        INTEGER,
        CURRENCY,
        NUMERIC,
        STRING
    }

    public DefaultCellGenerator(RowGenerator rowGenerator, int columnNum) {
        this.rowGenerator = rowGenerator;
        this.columnNum = columnNum;
        this.extendedCellStyle = new ExtendedCellStyle();
        cellType = CellType.STRING;
    }

    @Override
    public CellGenerator havingValue(Object cellValue) {
        this.cellValue = cellValue;
        return this;
    }

    @Override
    public CellGenerator usingStyle(String cellStyle) {
        this.extendedCellStyle = cellStyleManager().copy(cellStyle);
        return this;
    }

    @Override
    public CellGenerator withDateFormat() {
        this.extendedCellStyle.withDataFormat("m/d/yyyy");
        this.cellType = CellType.DATE;
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
        cellStyleManager().saveStyle(cellStyleName, this.extendedCellStyle.copy());
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

    public void applyCell(Row row) {
        Cell cell = row.createCell(columnNum);
        if (cellValue == null) {
            cell.setCellType(BLANK);
        } else {
            switch (cellType) {
                case DATE:
                    cell.setCellValue(getDateValue());
                    break;
                case INTEGER:
                case NUMERIC:
                case CURRENCY:
                    cell.setCellType(NUMERIC);
                    cell.setCellValue(getNumericValue());
                    break;
                default:
                    cell.setCellType(STRING);
                    cell.setCellValue(cellValue.toString());
                    break;
            }
        }
        CellStyle cellStyle = cellStyleManager().getCellStyle(row.getSheet().getWorkbook(), this.extendedCellStyle);
        cell.setCellStyle(cellStyle);
    }

    private double getNumericValue() {
        if (cellValue instanceof String) {
            String stringValue = (String) cellValue;
            return Double.valueOf(stringValue);
        }
        if (cellValue instanceof Number) {
            Number number = (Number) cellValue;
            return number.doubleValue();
        }
        return Double.NaN;
    }

    private Date getDateValue() {
        if (cellValue instanceof Date) {
            return (Date) cellValue;
        }
        if (cellValue instanceof String) {
            String stringValue = (String) cellValue;
            try {
                return new SimpleDateFormat("MM/dd/yyyy").parse(stringValue);
            } catch (ParseException e) {
                throw new RuntimeException(e);
            }
        }
        return null;
    }
}
