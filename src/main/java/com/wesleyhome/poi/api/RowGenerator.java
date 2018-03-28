package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;

import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public interface RowGenerator extends WorkbookCreator {
    default RowGenerator withStartColumn(String columnName){
        return withStartColumn(convertColStringToIndex(columnName));
    }

    RowGenerator withStartColumn(int columnNum);

    CellGenerator cell(int columnNum);
}
