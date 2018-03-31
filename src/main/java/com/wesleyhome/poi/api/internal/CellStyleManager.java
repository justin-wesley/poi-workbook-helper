package com.wesleyhome.poi.api.internal;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;

import static java.util.Comparator.comparingInt;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

public class CellStyleManager {

    private ExtendedMap<String, ExtendedCellStyle> styleMap;
    private ExtendedMap<Workbook, Map<ExtendedCellStyle, CellStyle>> cellStyles;

    CellStyleManager(){
        styleMap = new ExtendedTreeMap<>();
        cellStyles = new ExtendedTreeMap<>(comparingInt(Object::hashCode));
    }

    ExtendedCellStyle copy(String cellStyle) {
        return styleMap.getOrDefault(cellStyle, new ExtendedCellStyle()).copy();
    }

    CellStyle getCellStyle(Workbook workbook, ExtendedCellStyle cs) {
        CellStyle cellStyle = cellStyles.getOrDefault(workbook, HashMap::new).get(cs);
        if(cellStyle != null) {
            IndexedColors backgroundColor = cs.getBackgroundColor();
            IndexedColors bottomBorderColor = cs.getBottomBorderColor();
            BorderStyle bottomBorderStyle = cs.getBottomBorderStyle();
            String dataFormatString = cs.getDataFormatString();
            IndexedColors fontColor = cs.getFontColor();
            Integer fontHeightInPoints = cs.getFontHeightInPoints();
            String fontName = cs.getFontName();
            HorizontalAlignment horizontalAlignment = cs.getHorizontalAlignment();
            IndexedColors leftBorderColor = cs.getLeftBorderColor();
            BorderStyle leftBorderStyle = cs.getLeftBorderStyle();
            IndexedColors rightBorderColor = cs.getRightBorderColor();
            BorderStyle rightBorderStyle = cs.getRightBorderStyle();
            IndexedColors topBorderColor = cs.getTopBorderColor();
            BorderStyle topBorderStyle = cs.getTopBorderStyle();
            VerticalAlignment verticalAlignment = cs.getVerticalAlignment();
            boolean bold = cs.isBold();
            boolean italic = cs.isItalic();
            boolean wrappedText = cs.isWrappedText();
            cellStyle = workbook.createCellStyle();
            setProperty(cellStyle, backgroundColor, IndexedColors::getIndex, this::applyBackgroundColor);
            setProperty(bottomBorderColor, IndexedColors::getIndex, cellStyle::setBottomBorderColor);
            setProperty(bottomBorderStyle, cellStyle::setBorderBottom);
            setProperty(topBorderColor, IndexedColors::getIndex, cellStyle::setTopBorderColor);
            setProperty(topBorderStyle, cellStyle::setBorderTop);
            setProperty(rightBorderColor, IndexedColors::getIndex, cellStyle::setRightBorderColor);
            setProperty(rightBorderStyle, cellStyle::setBorderRight);
            setProperty(leftBorderColor, IndexedColors::getIndex, cellStyle::setLeftBorderColor);
            setProperty(leftBorderStyle, cellStyle::setBorderLeft);
            setProperty(verticalAlignment, cellStyle::setVerticalAlignment);
            setProperty(horizontalAlignment, cellStyle::setAlignment);
            setProperty(dataFormatString, fmt -> workbook.createDataFormat().getFormat(fmt), cellStyle::setDataFormat);
            cellStyle.setWrapText(wrappedText);
            if(fontColor != null || fontHeightInPoints != null || fontName != null || bold || italic){
                Font font = workbook.createFont();
                setProperty(fontColor, IndexedColors::getIndex, font::setColor);
                setProperty(fontHeightInPoints, Integer::shortValue, font::setFontHeightInPoints);
                setProperty(fontName, font::setFontName);
                font.setBold(bold);
                font.setItalic(italic);
                cellStyle.setFont(font);
            }
        }
        return cellStyle;
    }

    private void applyBackgroundColor(CellStyle cs1, Short index) {
        cs1.setFillForegroundColor(index);
        cs1.setFillPattern(SOLID_FOREGROUND);
    }

    private <T, U> void setProperty(T value, Function<T, U> supplier, Consumer<U> consumer){
        if (value != null) {
            U apply = supplier.apply(value);
            consumer.accept(apply);
        }
    }

    private <T, U, P> void setProperty(P style, T value, Function<T, U> supplier, BiConsumer<P, U> consumer){
        if (value != null) {
            U apply = supplier.apply(value);
            consumer.accept(style, apply);
        }
    }

    private <T> void setProperty(T value, Consumer<T> consumer){
        if (value != null) {
            consumer.accept(value);
        }
        consumer.accept(value);
    }

    public void saveStyle(String cellStyleName, ExtendedCellStyle cellStyle) {
        cellStyle.markImmutable();
        styleMap.put(cellStyleName, cellStyle);
    }
}
