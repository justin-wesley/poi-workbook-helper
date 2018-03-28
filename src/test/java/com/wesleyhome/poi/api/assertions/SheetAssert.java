package com.wesleyhome.poi.api.assertions;

import com.wesleyhome.poi.api.CellNameHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.assertj.core.api.AbstractAssert;

import static com.wesleyhome.poi.api.CellNameHelper.convertToCell;

public class SheetAssert extends AbstractAssert<SheetAssert, Sheet> {
    SheetAssert(Sheet sheet) {
        super(sheet, SheetAssert.class);
    }

    public CellAssert cell(String cellCoordinates) {
        isNotNull();
        Cell cell = convertToCell(actual, cellCoordinates);
        return new CellAssert(cell);
    }
}
