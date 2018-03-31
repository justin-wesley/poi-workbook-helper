package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public final class CellNameHelper {

    private static final String CELL_REGEX = "([A-Za-z]{1,2})(\\d+)";
    public static final Pattern CELL_PATTERN = Pattern.compile(CELL_REGEX);

    public static Cell convertToCell(Sheet sheet, String cellCoordinates){
        Matcher matcher = CELL_PATTERN.matcher(cellCoordinates);
        if(matcher.matches()){
            String columnName = matcher.group(1);
            int rowNum = Integer.parseInt(matcher.group(2))-1;
            int columnNum = convertColStringToIndex(columnName);
            return sheet.getRow(rowNum).getCell(columnNum);
        }
        throw new IllegalArgumentException(String.format("%s is not a valid Cell Reference Name", cellCoordinates));
    }
}
