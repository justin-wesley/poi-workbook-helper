package com.wesleyhome.poi.api;


import com.wesleyhome.poi.api.creator.WorkbookCreator;
import com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator;

import java.util.Iterator;
import java.util.function.*;

public interface WorkbookGenerator extends WorkbookCreator {

    WorkbookGenerator generateStyles(Consumer<CellStyler> cellStyler);

    default <T> WorkbookGenerator generateSheets(Iterable<T> iterable, BiConsumer<WorkbookGenerator, T> sheetGeneratorFunction) {
        Iterator<T> i = iterable.iterator();
        while(i.hasNext()) {
            T t = i.next();
            sheetGeneratorFunction.accept(this, t);
        }
        return this;
    }

    static WorkbookGenerator createNewWorkbook() {
        return createNewWorkbook(WorkbookType.EXCEL_OPEN);
    }

    static WorkbookGenerator createNewWorkbook(WorkbookType workbookType) {
        return new DefaultWorkbookGenerator(workbookType);
    }

    static SheetGenerator createSheet() {
        return createSheet(WorkbookType.EXCEL_OPEN);
    }

    static SheetGenerator createSheet(WorkbookType workbookType){
        return createSheet(workbookType, "Sheet1");
    }

    static SheetGenerator createSheet(String sheetName) {
        return createNewWorkbook().sheet(sheetName);
    }

    static SheetGenerator createSheet(WorkbookType workbookType, String sheetName) {
        return createNewWorkbook(workbookType).sheet(sheetName);
    }


}
