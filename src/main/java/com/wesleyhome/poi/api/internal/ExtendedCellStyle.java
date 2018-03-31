package com.wesleyhome.poi.api.internal;

import lombok.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import static com.wesleyhome.poi.api.internal.Immutable.immutable;
import static lombok.AccessLevel.PACKAGE;

/**
 * This will hold all the cell style information and can be used to manage cell styles for the workbook
 */
@Getter(PACKAGE)
@EqualsAndHashCode
@ToString
@Builder(toBuilder = true)
@NoArgsConstructor(access = PACKAGE)
class ExtendedCellStyle {

    private HorizontalAlignment horizontalAlignment;
    private BorderStyle topBorderStyle;
    private IndexedColors topBorderColor;
    private BorderStyle bottomBorderStyle;
    private IndexedColors bottomBorderColor;
    private BorderStyle leftBorderStyle;
    private IndexedColors leftBorderColor;
    private BorderStyle rightBorderStyle;
    private IndexedColors rightBorderColor;
    private String dataFormatString;
    private IndexedColors backgroundColor;
    private IndexedColors fontColor;
    private String fontName;
    private Integer fontHeightInPoints;
    private boolean italic;
    private boolean bold;
    private boolean wrappedText;
    private VerticalAlignment verticalAlignment;
    private boolean immutable;

    ExtendedCellStyle withHorizonalAlignment(HorizontalAlignment horizontalAlignment) {
        checkImmutable();
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    private void checkImmutable() {
        if (immutable) {
            throw new IllegalArgumentException("Cannot update an immutable style");
        }
    }

    private ExtendedCellStyle withWrappedText() {
        checkImmutable();
        this.wrappedText = true;
        return this;
    }

    private ExtendedCellStyle withVerticalAlignment(VerticalAlignment verticalAlignment) {
        checkImmutable();
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    ExtendedCellStyle withTopBorder(BorderStyle topBorderStyle, IndexedColors topBorderColor) {
        checkImmutable();
        this.topBorderStyle = topBorderStyle;
        this.topBorderColor = topBorderColor;
        return this;
    }

    ExtendedCellStyle withBottomBorder(BorderStyle bottomBorderStyle, IndexedColors bottomBorderColor) {
        checkImmutable();
        this.bottomBorderStyle = bottomBorderStyle;
        this.bottomBorderColor = bottomBorderColor;
        return this;
    }

    ExtendedCellStyle withLeftBorder(BorderStyle leftBorderStyle, IndexedColors leftBorderColor) {
        checkImmutable();
        this.leftBorderStyle = leftBorderStyle;
        this.leftBorderColor = leftBorderColor;
        return this;
    }

    ExtendedCellStyle withRightBorder(BorderStyle rightBorderStyle, IndexedColors rightBorderColor) {
        checkImmutable();
        this.rightBorderStyle = rightBorderStyle;
        this.rightBorderColor = rightBorderColor;
        return this;
    }

    ExtendedCellStyle withDataFormat(String dataFormat) {
        checkImmutable();
        this.dataFormatString = dataFormat;
        return this;
    }

    ExtendedCellStyle withBackgroundColor(IndexedColors backgroundColor) {
        checkImmutable();
        this.backgroundColor = backgroundColor;
        return this;
    }

    // Font Properties

    ExtendedCellStyle withFontColor(IndexedColors fontColor) {
        checkImmutable();
        this.fontColor = fontColor;
        return this;
    }

    private ExtendedCellStyle withItalic() {
        checkImmutable();
        this.italic = true;
        return this;
    }

    private ExtendedCellStyle withBold() {
        checkImmutable();
        this.bold = true;
        return this;
    }

    private ExtendedCellStyle withFontHeight(int fontHeightInPoints) {
        checkImmutable();
        this.fontHeightInPoints = fontHeightInPoints;
        return this;
    }

    private ExtendedCellStyle withFontName(String fontName) {
        checkImmutable();
        this.fontName = fontName;
        return this;
    }

    public ExtendedCellStyle copy() {
        ExtendedCellStyleBuilder builder = this.toBuilder();
        return builder.build();
    }

    void markImmutable() {
        this.immutable = true;
    }
}
