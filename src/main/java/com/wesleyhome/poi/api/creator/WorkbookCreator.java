package com.wesleyhome.poi.api.creator;

import com.wesleyhome.poi.api.*;
import com.wesleyhome.poi.api.internal.CellStyleManager;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.regex.Matcher;

import static java.lang.String.format;
import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public interface WorkbookCreator {

    Workbook createWorkbook();

    default SheetGenerator sheet(){
        return sheet("Sheet1");
    }

    SheetGenerator sheet(String sheetName);

    RowGenerator nextRow();

    RowGenerator row(int rowNum);

    CellGenerator nextCell();

    CellGenerator cell(int columnNumber, int rowNum);

    default CellGenerator cell(String cellName){
        Matcher matcher = CellNameHelper.CELL_PATTERN.matcher(cellName);
        if(matcher.matches()) {
            String columnName = matcher.group(1);
            int rowNum = Integer.parseInt(matcher.group(2)) - 1;
            int colNum = convertColStringToIndex(columnName);
            return cell(colNum, rowNum);
        }
        throw new IllegalArgumentException(format("%s is not a valid Cell Reference Name", cellName));
    }

    CellStyleManager cellStyleManager();
}
