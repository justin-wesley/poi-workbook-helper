package com.wesleyhome.poi.api.internal;

import lombok.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

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
        checkImmutable();
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    private void checkImmutable() {
        if (immutable) {
            throw new IllegalArgumentException("Cannot update an immutable style");
        }
    }

    ExtendedCellStyle withWrappedText() {
        checkImmutable();
        this.wrappedText = true;
        return this;
    }

    ExtendedCellStyle withoutWrappedText() {
        checkImmutable();
        this.wrappedText = false;
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

    ExtendedCellStyle withBold() {
        checkImmutable();
        this.bold = true;
        return this;
    }

    ExtendedCellStyle withoutBold() {
        checkImmutable();
        this.bold = false;
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

    public ExtendedCellStyle withDataFormat(int index) {
        this.dataFormat = index;
        return this;
    }

    void markImmutable() {
        this.immutable = true;
    }

    public ExtendedCellStyle copy() {
        return toBuilder()
            .immutable(false)
            .build();
    }

    public ExtendedCellStyle noBackgroundColor() {
        this.backgroundColor = null;
        return this;
    }

}
