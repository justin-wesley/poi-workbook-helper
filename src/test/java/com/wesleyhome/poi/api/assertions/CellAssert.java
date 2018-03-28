package com.wesleyhome.poi.api.assertions;

import com.wesleyhome.poi.api.CellGenerator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.assertj.core.api.AbstractAssert;
import org.assertj.core.util.Objects;

public class CellAssert extends AbstractAssert<CellAssert, Cell> {
    CellAssert(Cell cell) {
        super(cell, CellAssert.class);
    }

    public CellAssert hasValue(Object expectedValue) {
        isNotNull();
        CellType cellTypeEnum = actual.getCellTypeEnum();
        Object actualValue;
        switch (cellTypeEnum){
            case STRING:
                actualValue = actual.getStringCellValue();
                break;
            case BOOLEAN:
                actualValue = actual.getBooleanCellValue();
                break;
            case NUMERIC:
                actualValue = actual.getNumericCellValue();
                break;
            default:
                actualValue = null;
                break;
        }
        if(!Objects.areEqual(actualValue, expectedValue)){
            failWithMessage("Expected cell value <%s> but was <%s>", expectedValue, actualValue);
        }
        return myself;
    }

    public CellAssert cell(String cellCoordinates) {
        return new SheetAssert(actual.getSheet()).cell(cellCoordinates);
    }
}
