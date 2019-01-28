package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellStyler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;

import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;

import static java.util.Comparator.comparingInt;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

public class CellStyleManager {

    private Map<String, CellStyler> cellStylerMap;
    //    private ExtendedMap<String, ExtendedCellStyle> styleMap;
    private Map<Workbook, Map<ExtendedCellStyle, CellStyle>> cellStyles;

    CellStyleManager() {
//        styleMap = new ExtendedTreeMap<>();
        cellStylerMap = new HashMap<>();
        cellStyles = new HashMap<>();
    }

    CellStyler get(String cellStyle) {
        return cellStylerMap.computeIfAbsent(cellStyle, cs -> new DefaultCellStyler());
    }

    ExtendedCellStyle copy(String cellStyle) {
        return get(cellStyle).getCellStyle();
    }

    private void applyBackgroundColor(CellStyle cs1, Short index) {
        cs1.setFillForegroundColor(index);
        cs1.setFillPattern(SOLID_FOREGROUND);
    }

    private <T, U> void setProperty(T value, Function<T, U> supplier, Consumer<U> consumer) {
        if (value != null) {
            U apply = supplier.apply(value);
            consumer.accept(apply);
        }
    }

    private <T, U, P> void setProperty(P style, T value, Function<T, U> supplier, BiConsumer<P, U> consumer) {
        if (value != null) {
            U apply = supplier.apply(value);
            consumer.accept(style, apply);
        }
    }

    private <T> void setProperty(T value, Consumer<T> consumer) {
        if (value != null) {
            consumer.accept(value);
        }
    }

    public void saveStyle(String cellStyleName, CellStyler cellStyler) {
        this.cellStylerMap.put(cellStyleName, cellStyler.immutable());
    }

    public void applyCellStyle(CellRangeAddress cellRangeAddress, Sheet sheet, ExtendedCellStyle cs) {
        CellStyle cellStyle = getCellStyle(sheet.getWorkbook(), cs, false);
        Cell cell = CellUtil.getCell(CellUtil.getRow(cellRangeAddress.getFirstRow(), sheet), cellRangeAddress.getFirstColumn());
        cell.setCellStyle(cellStyle);
        if (cs.getBottomBorderColor() != null) {
            RegionUtil.setBottomBorderColor(cs.getBottomBorderColor().index, cellRangeAddress, sheet);
        }
        if (cs.getTopBorderColor() != null) {
            RegionUtil.setTopBorderColor(cs.getTopBorderColor().index, cellRangeAddress, sheet);
        }
        if (cs.getLeftBorderColor() != null) {
            RegionUtil.setLeftBorderColor(cs.getLeftBorderColor().index, cellRangeAddress, sheet);
        }
        if (cs.getRightBorderColor() != null) {
            RegionUtil.setRightBorderColor(cs.getRightBorderColor().index, cellRangeAddress, sheet);
        }
        RegionUtil.setBorderLeft(cs.getLeftBorderStyle(), cellRangeAddress, sheet);
        RegionUtil.setBorderRight(cs.getRightBorderStyle(), cellRangeAddress, sheet);
        RegionUtil.setBorderBottom(cs.getBottomBorderStyle(), cellRangeAddress, sheet);
        RegionUtil.setBorderTop(cs.getTopBorderStyle(), cellRangeAddress, sheet);
    }

    public void applyCellStyle(Cell cell, ExtendedCellStyle extendedCellStyle) {
        CellStyle cellStyle = getCellStyle(cell.getRow().getSheet().getWorkbook(), extendedCellStyle, true);
        cell.setCellStyle(cellStyle);
    }

    private CellStyle getCellStyle(Workbook workbook, ExtendedCellStyle ecs, boolean applyBorder) {
        Map<ExtendedCellStyle, CellStyle> extendCSMap = cellStyles.computeIfAbsent(workbook, wb -> new HashMap<>());
        CellStyle cs = extendCSMap.computeIfAbsent(ecs, cs1 -> {
            CellStyle cellStyle = workbook.createCellStyle();
            IndexedColors backgroundColor = ecs.getBackgroundColor();
            Short dataFormat = ecs.getDataFormat();
            if(dataFormat == null) {
                String dataFormatString = ecs.getDataFormatString();
                if(dataFormatString != null){
                    DataFormat dataFormat1 = workbook.createDataFormat();
                    dataFormat = dataFormat1.getFormat(dataFormatString);
                }
            }
            IndexedColors fontColor = ecs.getFontColor();
            Integer fontHeightInPoints = ecs.getFontHeightInPoints();
            String fontName = ecs.getFontName();
            HorizontalAlignment horizontalAlignment = ecs.getHorizontalAlignment();
            VerticalAlignment verticalAlignment = ecs.getVerticalAlignment();
            boolean bold = ecs.isBold();
            boolean italic = ecs.isItalic();
            boolean underline = ecs.isUnderline();
            boolean wrappedText = ecs.isWrappedText();
            setProperty(cellStyle, backgroundColor, IndexedColors::getIndex, this::applyBackgroundColor);
            setProperty(verticalAlignment, cellStyle::setVerticalAlignment);
            setProperty(horizontalAlignment, cellStyle::setAlignment);
            setProperty(dataFormat, cellStyle::setDataFormat);
            cellStyle.setWrapText(wrappedText);
            if (fontColor != null || fontHeightInPoints != null || fontName != null || bold || italic) {
                Font font = workbook.createFont();
                setProperty(fontColor, IndexedColors::getIndex, font::setColor);
                setProperty(fontHeightInPoints, Integer::shortValue, font::setFontHeightInPoints);
                setProperty(fontName, font::setFontName);
                font.setBold(bold);
                font.setItalic(italic);
                if(underline) {
                    font.setUnderline(Font.U_SINGLE);
                }
                cellStyle.setFont(font);
            }
            return cellStyle;
        });
        if (applyBorder) {
            IndexedColors bottomBorderColor = ecs.getBottomBorderColor();
            BorderStyle bottomBorderStyle = ecs.getBottomBorderStyle();
            IndexedColors leftBorderColor = ecs.getLeftBorderColor();
            BorderStyle leftBorderStyle = ecs.getLeftBorderStyle();
            IndexedColors rightBorderColor = ecs.getRightBorderColor();
            BorderStyle rightBorderStyle = ecs.getRightBorderStyle();
            IndexedColors topBorderColor = ecs.getTopBorderColor();
            BorderStyle topBorderStyle = ecs.getTopBorderStyle();
            setProperty(bottomBorderColor, IndexedColors::getIndex, cs::setBottomBorderColor);
            setProperty(bottomBorderStyle, cs::setBorderBottom);
            setProperty(topBorderColor, IndexedColors::getIndex, cs::setTopBorderColor);
            setProperty(topBorderStyle, cs::setBorderTop);
            setProperty(rightBorderColor, IndexedColors::getIndex, cs::setRightBorderColor);
            setProperty(rightBorderStyle, cs::setBorderRight);
            setProperty(leftBorderColor, IndexedColors::getIndex, cs::setLeftBorderColor);
            setProperty(leftBorderStyle, cs::setBorderLeft);
        }
        return cs;
    }

}
