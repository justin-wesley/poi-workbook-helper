package com.wesleyhome.poi.api;


import com.wesleyhome.poi.api.creator.WorkbookCreator;
import com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator;

public interface WorkbookGenerator extends WorkbookCreator {

    static SheetGenerator create(){
        return new DefaultWorkbookGenerator().sheet("Sheet1");
    }
}
