package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.internal.ExtendedCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.function.*;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public interface CellStyler {
    ExtendedCellStyle DEFAULT_STYLE = ExtendedCellStyle.builder().immutable(true).build();

    String name();

    CellStyler withHorizontalAlignment(HorizontalAlignment horizontalAlignment);

    CellStyler withDateTimeFormat();

    CellStyler withDateFormat();

    CellStyler withIntegerFormat();

    CellStyler withCurrencyFormat();

    CellStyler withNumericStyle();

    CellStyler withAccountingFormat();

    CellStyler withBackgroundColor(IndexedColors color);

    CellStyler withFontColor(IndexedColors color);

    CellStyler withFontSize(int fontSize);

    CellStyler withBoldFont();

    CellStyler notBold();

    CellStyler withItalicFont();

    CellStyler withNoBackgroundColor();

    CellStyler withAllBorders(BorderStyle borderStyle, IndexedColors color);

    CellStyler withTopBorder(BorderStyle borderStyle, IndexedColors color);

    CellStyler withBottomBorder(BorderStyle borderStyle, IndexedColors color);

    CellStyler withLeftBorder(BorderStyle borderStyle, IndexedColors color);

    CellStyler withRightBorder(BorderStyle borderStyle, IndexedColors color);

    CellStyler withWrappedText();

    CellStyler withoutWrappedText();

    CellStyler as(String cellStyleName);

    CellStyler usingStyle(String cellStyle);

    void applyCellStyle(CellRangeAddress cellRangeAddress, Sheet sheet, String cellStyleName);

    void applyCellStyle(Cell cell, String cellStyleName);

    default CellStyler applyIf(Predicate<CellStyler> filter, Consumer<CellStyler> consumer) {
        if(filter.test(this)){
            consumer.accept(this);
        }
        return this;
    }

    default CellStyler applyIf(boolean filter, Function<CellStyler, CellStyler> function) {
        if(filter) {
            return function.apply(this);
        }
        return this;
    }

    CellStyler reset();

    CellStyler immutable();

    ExtendedCellStyle getCellStyle();

//    CellStyler usingStyles(String firstStyle, String... otherStyles);

    CellStyler mergeWith(CellStyler that);

    CellStyler usingStyles(Iterable<String> styles);

    default <T> CellStyler applyEach(T[] values, BiConsumer<CellStyler, T> biConsumer) {
        Stream.of(values)
            .forEach(v-> biConsumer.accept(this.reset(), v));
        return this;
    }

    default <T> CellStyler applyEach(Iterable<T> values, BiConsumer<CellStyler, T> biConsumer) {
        StreamSupport.stream(values.spliterator(), false)
            .forEach(v-> biConsumer.accept(this.reset(), v));
        return this;
    }
}
