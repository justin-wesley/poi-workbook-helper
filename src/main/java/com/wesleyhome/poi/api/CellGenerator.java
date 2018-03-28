package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;
import org.apache.poi.ss.usermodel.IndexedColors;

public interface CellGenerator extends WorkbookCreator {
    CellGenerator havingValue(Object cellValue);

    CellGenerator usingStyle(String cellStyle);

    CellGenerator withDateFormat();

    CellGenerator withIntegerFormat();

    CellGenerator withCurrencyFormat();

    CellGenerator withNumericStyle();

    CellGenerator withBackgroundColor(IndexedColors color);

    CellGenerator withFontColor(IndexedColors color);

    CellGenerator isBold();

    CellGenerator as(String cellStyleName);

    CellGenerator withNoBackgroundColor();

}
