package com.wesleyhome.poi.api.creator;

import com.wesleyhome.poi.api.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.regex.Matcher;

import static java.lang.String.format;
import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public interface WorkbookCreator {

    Workbook createWorkbook();

    WorkbookGenerator workbook();

    default SheetGenerator sheet(){
        return sheet("Sheet1");
    }

    SheetGenerator sheet(String sheetName);

    RowGenerator nextRow();

    RowGenerator row();

    RowGenerator row(int rowNum);

    CellGenerator nextCell();

    CellGenerator cell();

    CellGenerator cell(int rowNum, int columnNumber);

    int rowNum();

    int columnNum();

    default CellGenerator cell(String cellName){
        Matcher matcher = CellNameHelper.CELL_PATTERN.matcher(cellName);
        if(matcher.matches()) {
            String columnName = matcher.group(1);
            int rowNum = Integer.parseInt(matcher.group(2)) - 1;
            int colNum = convertColStringToIndex(columnName);
            return cell(rowNum, colNum);
        }
        throw new IllegalArgumentException(format("%s is not a valid Cell Reference Name", cellName));
    }

    CellStyler cellStyler();

    static String formatCellRange(CellRangeAddress cellRangeAddress) {
        CellAddress first = new CellAddress(cellRangeAddress.getFirstRow(), cellRangeAddress.getFirstColumn());
        CellAddress last = new CellAddress(cellRangeAddress.getLastRow(), cellRangeAddress.getLastColumn());
        return String.format("%s:%s", first.formatAsString(), last.formatAsString());
    }
}
