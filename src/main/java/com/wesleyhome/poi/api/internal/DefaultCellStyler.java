package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellStyler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.function.Function;
import java.util.stream.StreamSupport;

public class DefaultCellStyler implements CellStyler {

    private static final CellStyleManager cellStyleManager = new CellStyleManager();
    public static final String DEFAULT_STYLE_NAME = "DEFAULT_STYLE_NAME";
    public static final String CURRENCY_FORMAT = "\"$\"#,##0_);(\"$\"#,##0)";
    public static final String DATE_FORMAT = "m/d/yy";
    public static final String DATE_TIME_FORMAT = "m/d/yy h:mm AM/PM";
    public static final String NUMERIC_FORMAT = "#,##0.00";
    public static final String ACCOUNTING_FORMAT = "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
    public static final String ZIP_FORMAT = "00000-0000";
    private ExtendedCellStyle currentCellStyle;
    private String styleName;

    public DefaultCellStyler() {
        currentCellStyle = DEFAULT_STYLE.copy();
        this.styleName = DEFAULT_STYLE_NAME;
    }

    private DefaultCellStyler(String newName, ExtendedCellStyle cellStyle) {
        this.styleName = newName;
        this.currentCellStyle = cellStyle;
    }

    @Override
    public String name() {
        return this.styleName;
    }

    public DefaultCellStyler(String styleName) {
        this.styleName = styleName;
        this.currentCellStyle = cellStyleManager.copy(styleName).copy();
    }

    @Override
    public CellStyler reset() {
        return new DefaultCellStyler();
    }

    @Override
    public CellStyler immutable() {
        this.currentCellStyle.markImmutable();
        return this;
    }

    @Override
    public CellStyler withHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        return applyCellStyle(ecs -> ecs.withHorizontalAlignment(horizontalAlignment));
    }

    @Override
    public CellStyler withDateTimeFormat() {
        return applyFormat(DATE_TIME_FORMAT);
    }

    @Override
    public CellStyler withDateFormat() {
        return applyFormat(DATE_FORMAT);
    }

    @Override
    public CellStyler withIntegerFormat() {
        return applyCellStyle(ecs -> ecs.withDataFormat(1));
    }

    @Override
    public CellStyler withCurrencyFormat() {
        return applyFormat(CURRENCY_FORMAT);
    }

    @Override
    public CellStyler withNumericFormat() {
        return applyFormat(NUMERIC_FORMAT);
    }

    @Override
    public CellStyler withZipCodeFormat() {
        return applyFormat(ZIP_FORMAT);
    }

    @Override
    public CellStyler withAccountingFormat() {
        return applyFormat(ACCOUNTING_FORMAT);
    }

    @Override
    public CellStyler withBackgroundColor(IndexedColors color) {
        return applyCellStyle(ecs -> ecs.withBackgroundColor(color));
    }

    @Override
    public CellStyler withFontColor(IndexedColors color) {
        return applyCellStyle(ecs -> ecs.withFontColor(color));
    }

    @Override
    public CellStyler withFontSize(int fontHeightInPoints) {
        return applyCellStyle(ecs -> ecs.withFontSize(fontHeightInPoints));
    }

    @Override
    public CellStyler withBoldFont() {
        return applyCellStyle(ExtendedCellStyle::withBoldFont);
    }

    @Override
    public CellStyler withItalicFont() {
        return applyCellStyle(ExtendedCellStyle::withItalicFont);
    }

    @Override
    public CellStyler notBold() {
        return applyCellStyle(ExtendedCellStyle::withoutBold);
    }

    @Override
    public CellStyler withNoBackgroundColor() {
        return applyCellStyle(ExtendedCellStyle::noBackgroundColor);
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
        return applyCellStyle(ecs -> ecs.withTopBorder(borderStyle, color));
    }

    @Override
    public CellStyler withBottomBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs -> ecs.withBottomBorder(borderStyle, color));
    }

    @Override
    public CellStyler withLeftBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs -> ecs.withLeftBorder(borderStyle, color));
    }

    @Override
    public CellStyler withRightBorder(BorderStyle borderStyle, IndexedColors color) {
        return applyCellStyle(ecs -> ecs.withRightBorder(borderStyle, color));
    }

    @Override
    public CellStyler withWrappedText() {
        return applyCellStyle(ExtendedCellStyle::withWrappedText);
    }

    @Override
    public CellStyler withoutWrappedText() {
        return applyCellStyle(ExtendedCellStyle::withoutWrappedText);
    }

    private CellStyler applyFormat(String fmt) {

        return applyCellStyle(ecs -> {
            int builtinFormat = BuiltinFormats.getBuiltinFormat(fmt);
            return builtinFormat < 0 ? ecs.withDataFormatString(fmt) : ecs.withDataFormat(builtinFormat);
        });
    }

    private CellStyler applyCellStyle(Function<ExtendedCellStyle, ExtendedCellStyle> consumer) {
        if (currentCellStyle.isImmutable()) {
            DefaultCellStyler defaultCellStyler = new DefaultCellStyler(this.styleName);
            return defaultCellStyler.applyCellStyle(consumer);
        }
        this.currentCellStyle = consumer.apply(this.currentCellStyle);
        return this;
    }

    @Override
    public ExtendedCellStyle getCellStyle() {
        return this.currentCellStyle;
    }

    @Override
    public CellStyler as(String cellStyleName) {
        this.styleName = getNewName(cellStyleName);
        cellStyleManager.saveStyle(this.styleName, this);
        return this;
    }

    public String getNewName(String cellStyleName) {
        return DEFAULT_STYLE_NAME.equals(this.styleName) ?
            cellStyleName :
            cellStyleName + "_" + this.styleName;
    }

    @Override
    public void applyCellStyle(CellRangeAddress cellRangeAddress, Sheet sheet, String cellStyleName) {
        cellStyleManager.applyCellStyle(cellRangeAddress, sheet, usingStyle(cellStyleName).getCellStyle());
    }

    @Override
    public void applyCellStyle(Cell cell, String cellStyleName) {
        cellStyleManager.applyCellStyle(cell, usingStyle(cellStyleName).getCellStyle());
    }

    @Override
    public CellStyler usingStyles(Iterable<String> styles) {
        return StreamSupport.stream(styles.spliterator(), false)
            .map(this::usingStyle)
            .reduce(CellStyler::mergeWith)
            .orElseGet(DefaultCellStyler::new);
    }

    @Override
    public CellStyler usingStyle(String cellStyle) {
        return cellStyleManager.get(cellStyle);
    }

    @Override
    public CellStyler mergeWith(CellStyler that) {
        String newName = getNewName(that.name());
        ExtendedCellStyle cellStyle = this.getCellStyle().merge(that.getCellStyle());
        DefaultCellStyler styler = new DefaultCellStyler(newName, cellStyle);
        cellStyleManager.saveStyle(newName, styler);
        return styler;
    }
}
