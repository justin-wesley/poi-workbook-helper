package com.wesleyhome.poi.api.internal;

import lombok.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;

import static lombok.AccessLevel.PACKAGE;

/**
 * This will hold all the cell style information and can be used to manage cell styles for the workbook
 */
@Getter
@AllArgsConstructor(access = PACKAGE)
@NoArgsConstructor(access = PACKAGE)
@EqualsAndHashCode
@Builder(toBuilder = true)
@ToString
public class ExtendedCellStyle {

    private HorizontalAlignment horizontalAlignment;
    private BorderStyle topBorderStyle;
    private IndexedColors topBorderColor;
    private BorderStyle bottomBorderStyle;
    private IndexedColors bottomBorderColor;
    private BorderStyle leftBorderStyle;
    private IndexedColors leftBorderColor;
    private BorderStyle rightBorderStyle;
    private IndexedColors rightBorderColor;
    private Integer dataFormat;
    private IndexedColors backgroundColor;
    private IndexedColors fontColor;
    private String fontName;
    private Integer fontHeightInPoints;
    private boolean italic;
    private boolean bold;
    private boolean wrappedText;
    private VerticalAlignment verticalAlignment;
    private boolean immutable;

    ExtendedCellStyle withHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        return checkImmutable(t -> this.horizontalAlignment = horizontalAlignment);
    }

    private ExtendedCellStyle checkImmutable(Consumer<ExtendedCellStyle> consumer) {
        if (immutable) {
            throw new IllegalArgumentException("Cannot update an immutable style");
        }
        consumer.accept(this);
        return this;
    }

    ExtendedCellStyle withWrappedText() {
        return checkImmutable(t -> t.wrappedText = true);
    }

    ExtendedCellStyle withoutWrappedText() {
        return checkImmutable(t -> t.wrappedText = false);
    }

    private ExtendedCellStyle withVerticalAlignment(VerticalAlignment verticalAlignment) {
        return checkImmutable(t -> t.verticalAlignment = verticalAlignment);
    }

    ExtendedCellStyle withTopBorder(BorderStyle topBorderStyle, IndexedColors topBorderColor) {
        return checkImmutable(t -> {
            t.topBorderStyle = topBorderStyle;
            t.topBorderColor = topBorderColor;
        });
    }

    ExtendedCellStyle withBottomBorder(BorderStyle bottomBorderStyle, IndexedColors bottomBorderColor) {
        return checkImmutable(t -> {
            t.bottomBorderStyle = bottomBorderStyle;
            t.bottomBorderColor = bottomBorderColor;
        });
    }

    ExtendedCellStyle withLeftBorder(BorderStyle leftBorderStyle, IndexedColors leftBorderColor) {
        return checkImmutable(t -> {
            t.leftBorderStyle = leftBorderStyle;
            t.leftBorderColor = leftBorderColor;
        });
    }

    ExtendedCellStyle withRightBorder(BorderStyle rightBorderStyle, IndexedColors rightBorderColor) {
        return checkImmutable(t -> {
            t.rightBorderStyle = rightBorderStyle;
            t.rightBorderColor = rightBorderColor;
        });
    }

    ExtendedCellStyle withBackgroundColor(IndexedColors backgroundColor) {
        return checkImmutable(t -> t.backgroundColor = backgroundColor);
    }

    public ExtendedCellStyle noBackgroundColor() {
        return checkImmutable(t -> t.backgroundColor = null);
    }


    // Font Properties
    ExtendedCellStyle withFontColor(IndexedColors fontColor) {
        return checkImmutable(t -> t.fontColor = fontColor);
    }

    ExtendedCellStyle withFontSize(int fontHeightInPoints) {
        return checkImmutable(t->t.fontHeightInPoints = fontHeightInPoints);
    }

    ExtendedCellStyle withBoldFont() {
        return checkImmutable(t -> t.bold = true);
    }

    public ExtendedCellStyle withItalicFont() {
        return checkImmutable(t -> t.italic = true);
    }

    ExtendedCellStyle withoutBold() {
        return checkImmutable(t -> t.bold = false);
    }

    ExtendedCellStyle withFontName(String fontName) {
        return checkImmutable(t -> t.fontName = fontName);
    }

    public ExtendedCellStyle withDataFormat(int index) {
        return checkImmutable(t -> t.dataFormat = index);
    }

    void markImmutable() {
        this.immutable = true;
    }

    public ExtendedCellStyle copy() {
        return toBuilder()
            .immutable(false)
            .build();
    }

    public ExtendedCellStyle merge(ExtendedCellStyle c) {
        ExtendedCellStyleBuilder b = this.toBuilder();
        applyIf(b::backgroundColor, c::getBackgroundColor)
            .applyIf(b::bold, c::isBold)
            .applyIf(b::bottomBorderColor, c::getBottomBorderColor)
            .applyIf(b::bottomBorderStyle, c::getBottomBorderStyle)
            .applyIf(b::dataFormat, c::getDataFormat)
            .applyIf(b::fontColor, c::getFontColor)
            .applyIf(b::fontHeightInPoints, c::getFontHeightInPoints)
            .applyIf(b::fontName, c::getFontName)
            .applyIf(b::horizontalAlignment, c::getHorizontalAlignment)
            .applyIf(b::italic, c::isItalic)
            .applyIf(b::leftBorderColor, c::getLeftBorderColor)
            .applyIf(b::leftBorderStyle, c::getLeftBorderStyle)
            .applyIf(b::rightBorderColor, c::getRightBorderColor)
            .applyIf(b::rightBorderStyle, c::getRightBorderStyle)
            .applyIf(b::topBorderColor, c::getTopBorderColor)
            .applyIf(b::topBorderStyle, c::getTopBorderStyle)
            .applyIf(b::verticalAlignment, c::getVerticalAlignment)
            .applyIf(b::wrappedText, c::isWrappedText);
        return b.build();
    }

    private <T> ExtendedCellStyle applyIf(Function<T, ExtendedCellStyleBuilder> consumer, Supplier<T> supplier) {
        if (supplier.get() != null) {
            consumer.apply(supplier.get());
        }
        return this;
    }
}
