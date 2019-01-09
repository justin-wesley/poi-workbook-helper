package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.*;
import java.util.Date;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static com.wesleyhome.poi.api.creator.WorkbookCreator.formatCellRange;
import static com.wesleyhome.poi.api.internal.DefaultCellGenerator.GeneratorCellType.*;
import static org.apache.poi.ss.usermodel.CellType.BLANK;

public class DefaultCellGenerator implements CellGenerator, Comparable<DefaultCellGenerator> {
    private final RowGenerator rowGenerator;
    private final int columnNum;
    private CellStyler cellStyler;
    private int cellsToMerge = 0;
    //    private ExtendedCellStyle extendedCellStyle;
    private Object cellValue;
    private GeneratorCellType cellType;
    private Float cellWidth;
    private String cellStyleName;

    enum GeneratorCellType {
        DATE,
        BOOLEAN,
        INTEGER,
        //        CURRENCY,
        NUMERIC,
        STRING,
        FORMULA
    }

    public DefaultCellGenerator(RowGenerator rowGenerator, int columnNum) {
        this.rowGenerator = rowGenerator;
        this.columnNum = columnNum;
        int rowNum = rowGenerator.rowNum();
        this.cellStyleName = new CellAddress(rowNum, columnNum).formatAsString();
        cellStyler = new DefaultCellStyler();
    }

    @Override
    public int columnNum() {
        return columnNum;
    }

    @Override
    public CellAddress cellAddress() {
        return new CellAddress(this.rowNum(), this.columnNum());
    }

    @Override
    public int rowNum() {
        return rowGenerator.rowNum();
    }

    @Override
    public SheetGenerator sheet() {
        return rowGenerator.sheet();
    }

    @Override
    public CellGenerator autosize() {
        return this.row().sheet().autosize(this.columnNum).cell();
    }

    @Override
    public CellGenerator hide() {
        return this.row().sheet().hide(this.columnNum).cell();
    }

    @Override
    public CellGenerator havingValue(Object cellValue) {
        this.cellValue = cellValue;
        if (hasValue() && this.cellType != null) {
            switch (cellType) {
                case DATE:
                    checkCellType(cellValue.getClass(), (cellValue instanceof Date || cellValue instanceof LocalDate || cellValue instanceof LocalDateTime));
                    break;
                case INTEGER:
                    checkCellType(cellValue.getClass(), cellValue instanceof Integer);
                    break;
                case NUMERIC:
                    checkCellType(cellValue.getClass(), cellValue instanceof Number);
                    break;
                case BOOLEAN:
                    checkCellType(cellValue.getClass(), cellValue instanceof Boolean || cellValue instanceof String);
                    break;
                case FORMULA:
                case STRING:
                    break;
            }
        } else if (hasValue() && this.cellType == null) {
            if (cellValue instanceof Integer) {
                this.cellType = INTEGER;
            } else if (cellValue instanceof Number) {
                this.cellType = NUMERIC;
            } else if (cellValue instanceof Date || cellValue instanceof LocalDate || cellValue instanceof LocalDateTime) {
                this.cellType = DATE;
            } else if (cellValue instanceof Boolean) {
                this.cellType = BOOLEAN;
            } else {
                this.cellType = GeneratorCellType.STRING;
            }
        }
        return this;
    }

    public void checkCellType(Class<?> cellValueClass, boolean validState) {
        if (!validState) {
            throw new IllegalArgumentException(String.format("Cell value has been assigned a cell type and value does not match type. %s != %s", cellType, cellValueClass));
        }
    }

    @Override
    public CellGenerator sumOfPreviousXRows(int numberOfRows) {

        int teamRowNum = this.rowGenerator.rowNum();
        int columnNum = this.columnNum;
        int startRow = teamRowNum - numberOfRows;
        int endRow = teamRowNum - 1;
        CellRangeAddress cellAddress = cell(startRow, columnNum).rangeAddress(endRow, columnNum);
        String cellRange = formatCellRange(cellAddress);
        return withFormula("SUM(%s)", cellRange);
    }

    @Override
    public CellRangeAddress rangeAddress(int endRow, int endColumn) {
        return new CellRangeAddress(rowNum(), endRow, columnNum(), endColumn);
    }

    @Override
    public CellGenerator withFormula(String formulaTemplate, Object... args) {
        this.cellValue = String.format(formulaTemplate, args);
        this.cellType = GeneratorCellType.FORMULA;
        return this;
    }

    @Override
    public CellGenerator mergeWithNextXCells(int numberOfCellsToMerge) {
        this.cellsToMerge = numberOfCellsToMerge;
        return this;
    }

    int getCellsToMerge() {
        return cellsToMerge;
    }

    @Override
    public CellStyler withHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        return this.cellStyler.withHorizontalAlignment(horizontalAlignment);
    }

    @Override
    public CellGenerator usingStyles(Iterable<String> styles) {
        this.cellStyler = cellStyler().usingStyles(styles);
        this.cellStyleName = this.cellStyler.name();
        return this;
    }

    @Override
    public CellGenerator usingStyles(String firstStyle, String... otherStyles) {
        return usingStyles(Stream.concat(Stream.of(firstStyle), Stream.of(otherStyles)).collect(Collectors.toList()));
    }

    @Override
    public CellGenerator usingStyle(String cellStyle) {
        this.cellStyleName = cellStyle;
        this.cellStyler = cellStyler().usingStyle(cellStyle);
        return this;
    }

    @Override
    public CellGenerator withDateFormat() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.DATE);
    }

    @Override
    public CellGenerator withDateTimeFormat() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.DATE_TIME);
    }

    @Override
    public CellGenerator withIntegerFormat() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.INTEGER);
    }

    @Override
    public CellGenerator withAccountingFormat() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.ACCOUNTING);
    }

    @Override
    public CellGenerator withCurrencyFormat() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.CURRENCY);
    }

    @Override
    public CellGenerator withNumericStyle() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.NUMERIC);
    }


    @Override
    public CellStyler withAllBorders(BorderStyle borderStyle, IndexedColors color) {
        return this.withTopBorder(borderStyle, color)
            .withBottomBorder(borderStyle, color)
            .withLeftBorder(borderStyle, color)
            .withRightBorder(borderStyle, color);
    }

    @Override
    public CellStyler withTopBorder(BorderStyle borderStyle, IndexedColors color) {
        return this.cellStyler.withTopBorder(borderStyle, color);
    }

    @Override
    public CellStyler withBottomBorder(BorderStyle borderStyle, IndexedColors color) {
        return this.cellStyler.withBottomBorder(borderStyle, color);
    }

    @Override
    public CellStyler withLeftBorder(BorderStyle borderStyle, IndexedColors color) {
        return this.cellStyler.withLeftBorder(borderStyle, color);
    }

    @Override
    public CellStyler withRightBorder(BorderStyle borderStyle, IndexedColors color) {
        return this.cellStyler.withRightBorder(borderStyle, color);
    }

    @Override
    public CellStyler withBackgroundColor(IndexedColors color) {
        return this.cellStyler.withBackgroundColor(color);
    }

    @Override
    public CellStyler withFontColor(IndexedColors color) {
        return this.cellStyler.withFontColor(color);
    }

    @Override
    public CellStyler isBold() {
        return this.cellStyler.isBold();
    }

    @Override
    public CellStyler notBold() {
        return this.cellStyler.notBold();
    }

    @Override
    public CellGenerator withNoBackgroundColor() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.NO_BACKGROUND_COLOR);
    }

    @Override
    public CellGenerator cellWidth(float cellWidth) {
        this.cellWidth = cellWidth;
        return this;
    }

    @Override
    public CellGenerator as(String cellStyleName) {
        this.cellStyleName = cellStyleName;
        this.cellStyler = this.cellStyler.as(cellStyleName);
        return this;
    }

    @Override
    public Workbook createWorkbook() {
        return this.rowGenerator.createWorkbook();
    }

    @Override
    public WorkbookGenerator workbook() {
        return this.rowGenerator.workbook();
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
    public RowGenerator row() {
        return rowGenerator;
    }

    @Override
    public CellGenerator cell() {
        return this;
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
    public CellGenerator withWrappedText() {
        return usingStyle(DefaultWorkbookGenerator.BuiltinStyles.WRAPPED_TEXT);
    }

    @Override
    public CellGenerator withoutWrappedText() {
        this.cellStyler = this.cellStyler.withoutWrappedText();
        return this;
    }

    @Override
    public CellGenerator cell(int rowNum, int columnNumber) {
        return this.rowGenerator.cell(rowNum, columnNumber);
    }

    @Override
    public CellStyler cellStyler() {
        return this.rowGenerator.cellStyler();
    }

    public void applyCell(Row row) {
        Cell cell = CellUtil.getCell(row, columnNum);
        if (!hasValue()) {
            cell.setCellType(BLANK);
        } else {
            switch (cellType) {
                case DATE:
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(getDateValue());
                    break;
                case INTEGER:
                case NUMERIC:
                    cell.setCellType(CellType.NUMERIC);
                    cell.setCellValue(getNumericValue());
                    break;
                case FORMULA:
                    cell.setCellType(CellType.FORMULA);
                    cell.setCellFormula(this.cellValue.toString());
                    break;
                case BOOLEAN:
                    cell.setCellType(CellType.BOOLEAN);
                    cell.setCellValue(getBooleanValue());
                    break;
                default:
                    cell.setCellType(CellType.STRING);
                    cell.setCellValue(cellValue.toString());
                    break;
            }
        }
        if (cellWidth != null) {
            int v = (int) (cellWidth * 256);
            cell.getRow().getSheet().setColumnWidth(columnNum, v);
        }
        if (cellsToMerge > 0) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(this.rowNum(), this.rowNum(), this.columnNum, this.columnNum + this.cellsToMerge);
            Sheet sheet = cell.getSheet();
            sheet.addMergedRegion(cellRangeAddress);
            this.cellStyler.applyCellStyle(cellRangeAddress, sheet, this.cellStyleName);
        } else {
            this.cellStyler.applyCellStyle(cell, this.cellStyleName);
        }
    }

    @Override
    public boolean hasValue() {
        return cellValue != null;
    }

    private boolean getBooleanValue() {
        if(cellValue instanceof Boolean) {
            return (boolean)cellValue;
        }
        if(cellValue instanceof String) {
            return Boolean.parseBoolean((String) cellValue);
        }
        return false;
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
        if (cellValue instanceof LocalDate) {
            LocalDate localDate = (LocalDate) cellValue;
            return Date.from(localDate.atStartOfDay().toInstant(ZoneOffset.ofHours(-6)));
        }
        if (cellValue instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) cellValue;
            return Date.from(Instant.from(localDateTime));
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

    @Override
    public int compareTo(DefaultCellGenerator o) {
        return Integer.compare(this.columnNum, o.columnNum);
    }

    @Override
    public String toString() {
        return String.format("[%s]->%s", this.cellAddress().formatAsString(), Objects.toString(this.cellValue, ""));
    }
}
