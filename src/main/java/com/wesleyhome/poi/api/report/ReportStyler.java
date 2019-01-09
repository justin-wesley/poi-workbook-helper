package com.wesleyhome.poi.api.report;

import org.apache.poi.ss.usermodel.IndexedColors;

public interface ReportStyler {
    default IndexedColors headerBackgroundColor() {
        return IndexedColors.DARK_BLUE;
    }

    default IndexedColors headerFontColor() {
        return IndexedColors.WHITE;
    }

    default boolean isHeaderBold() {
        return true;
    }

    default IndexedColors oddRowBackgroundColor() {
        return IndexedColors.WHITE;
    }

    default IndexedColors oddRowFontColor() {
        return IndexedColors.BLACK;
    }

    default boolean isEvenRowBold(){
        return false;
    }

    default IndexedColors evenRowBackgroundColor() {
        return IndexedColors.LIGHT_CORNFLOWER_BLUE;
    }

    default IndexedColors evenRowFontColor() {
        return IndexedColors.BLACK;
    }

    default boolean isOddRowBold(){
        return false;
    }

}
