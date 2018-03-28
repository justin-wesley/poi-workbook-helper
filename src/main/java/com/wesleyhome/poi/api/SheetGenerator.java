package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;

public interface SheetGenerator extends WorkbookCreator {

    SheetGenerator withSheetName(String test_sheet);

    /**
     *
     * @param rowNum (0-based)
     * @return
     */
    SheetGenerator withStartRow(int rowNum);
}
