package com.wesleyhome.poi.api;


import com.wesleyhome.poi.api.creator.WorkbookCreator;
import com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator;

public interface WorkbookGenerator extends WorkbookCreator {

    static SheetGenerator create() {
        return create(WorkbookType.EXCEL_OPEN);
    }

    static SheetGenerator create(WorkbookType workbookType){
        return create(workbookType, "Sheet1");
    }

    static SheetGenerator create(String sheetName) {
        return create(WorkbookType.EXCEL_OPEN, sheetName);
    }

    static SheetGenerator create(WorkbookType workbookType, String sheetName) {
        return new DefaultWorkbookGenerator(workbookType).sheet(sheetName);
    }
}
