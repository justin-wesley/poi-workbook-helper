package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Objects;
import java.util.function.Function;

import static com.wesleyhome.poi.api.creator.WorkbookCreator.formatCellRange;
import static org.apache.poi.ss.usermodel.CellType.*;

public class DefaultCellGenerator implements CellGenerator, Comparable<DefaultCellGenerator> {
    private final RowGenerator rowGenerator;
    private final int columnNum;
    private int cellsToMerge = 0;
    private ExtendedCellStyle extendedCellStyle;
    private Object cellValue;
    private GeneratorCellType cellType;
    private Float cellWidth;

    enum GeneratorCellType {
        DATE,
        INTEGER,
        CURRENCY,
        NUMERIC,
        STRING,
        FORMULA
    }

    public DefaultCellGenerator(RowGenerator rowGenerator, int columnNum) {
        this.rowGenerator = rowGenerator;
        this.columnNum = columnNum;
        this.extendedCellStyle = new ExtendedCellStyle();
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
    public CellGenerator withHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        return applyCellStyle(ecs->ecs.withHorizontalAlignment(horizontalAlignment));
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
        if (hasValue() && this.cellType == null) {
            if (cellValue instanceof Integer) {
                this.cellType = GeneratorCellType.INTEGER;
            }else if (cellValue instanceof Number) {
                this.cellType = GeneratorCellType.NUMERIC;
            } else if (cellValue instanceof Date || cellValue instanceof LocalDate || cellValue instanceof LocalDateTime) {
                this.cellType = GeneratorCellType.DATE;
            } else {
                this.cellType = GeneratorCellType.STRING;
            }
        }
        return this;
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
    public CellGenerator usingStyle(String cellStyle) {
        this.extendedCellStyle = cellStyleManager().copy(cellStyle);
        return this;
    }

    @Override
    public CellGenerator withDateFormat() {
        this.cellType = GeneratorCellType.DATE;
        return applyFormat("m/d/yy h:mm");
    }

    @Override
    public CellGenerator withIntegerFormat() {
        this.cellType = GeneratorCellType.INTEGER;
        return applyCellStyle(ecs -> ecs.withDataFormat(1));
    }

    @Override
    public CellGenerator withAccountingFormat() {
        return applyFormat("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)");
    }

    @Override
    public CellGenerator withCurrencyFormat() {
        return applyFormat("\"$\"#,##0_);(\"$\"#,##0)");
    }

    @Override
    public CellGenerator withNumericStyle() {
        return applyFormat("#,##0.00");
    }

    private CellGenerator applyFormat(String fmt) {
        int index = BuiltinFormats.getBuiltinFormat(fmt);
        return applyCellStyle(ecs -> ecs.withDataFormat(index));
    }

    @Override
    public CellGenerator withAllBorders(BorderStyle borderStyle, IndexedColors color) {
        return this.withTopBorder(borderStyle, color)
            .withBottomBorder(borderStyle, color)
            .withLeftBorder(borderStyle, color)
            .withRightBorder(borderStyle, color);
    }

    @Override
    public CellGenerator withTopBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs->ecs.withTopBorder(borderStyle, color));
    }

    @Override
    public CellGenerator withBottomBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs->ecs.withBottomBorder(borderStyle, color));
    }

    @Override
    public CellGenerator withLeftBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs->ecs.withLeftBorder(borderStyle, color));
    }

    @Override
    public CellGenerator withRightBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs->ecs.withRightBorder(borderStyle, color));
    }

    @Override
    public CellGenerator withBackgroundColor(IndexedColors color) {
        return applyCellStyle(ecs->ecs.withBackgroundColor(color));
    }

    @Override
    public CellGenerator withFontColor(IndexedColors color) {
        return applyCellStyle(ecs->ecs.withFontColor(color));
    }

    private CellGenerator applyCellStyle(Function<ExtendedCellStyle, ExtendedCellStyle> consumer) {
        this.extendedCellStyle = consumer.apply(this.extendedCellStyle);
        return this;
    }

    @Override
    public CellGenerator isBold() {
        return applyCellStyle(ExtendedCellStyle::withBold);
    }

    @Override
    public CellGenerator notBold() {
        return applyCellStyle(ExtendedCellStyle::withoutBold);
    }

    @Override
    public CellGenerator withNoBackgroundColor() {
        return applyCellStyle(ExtendedCellStyle::noBackgroundColor);
    }

    @Override
    public CellGenerator cellWidth(float cellWidth) {
        this.cellWidth = cellWidth;
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
        return applyCellStyle(ExtendedCellStyle::withWrappedText);
    }

    @Override
    public CellGenerator withoutWrappedText() {
        return applyCellStyle(ExtendedCellStyle::withoutWrappedText);
    }

    @Override
    public CellGenerator cell(int rowNum, int columnNumber) {
        return this.rowGenerator.cell(rowNum, columnNumber);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return this.rowGenerator.cellStyleManager();
    }

    public void applyCell(Row row) {
        Cell cell = CellUtil.getCell(row, columnNum);
        if (!hasValue()) {
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
                case FORMULA:
                    cell.setCellType(CellType.FORMULA);
                    cell.setCellFormula(this.cellValue.toString());
                    break;
                default:
                    cell.setCellType(STRING);
                    cell.setCellValue(cellValue.toString());
                    break;
            }
        }
        if(cellWidth != null) {
            int v = (int)(cellWidth * 256);
            cell.getRow().getSheet().setColumnWidth(columnNum, v);
        }
        if(cellsToMerge > 0) {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(this.rowNum(), this.rowNum(), this.columnNum, this.columnNum + this.cellsToMerge);
            Sheet sheet = cell.getSheet();
            sheet.addMergedRegion(cellRangeAddress);
            cellStyleManager().applyCellStyle(cellRangeAddress, sheet, this.extendedCellStyle);
        } else {
            cellStyleManager().applyCellStyle(cell, this.extendedCellStyle);
        }
    }

    @Override
    public boolean hasValue() {
        return cellValue != null;
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

    @Override
    public int compareTo(DefaultCellGenerator o) {
        return Integer.compare(this.columnNum, o.columnNum);
    }

    @Override
    public String toString() {
        return String.format("[%s]->%s", this.cellAddress().formatAsString(), Objects.toString(this.cellValue, ""));
    }
}
