package com.wesleyhome.poi.api.assertions;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.assertj.core.api.AbstractAssert;
import org.assertj.core.api.Assert;

import java.util.function.Supplier;

public class WorkbookAssert extends AbstractAssert<WorkbookAssert, Workbook> {
    private WorkbookAssert(Workbook actual) {
        super(actual, WorkbookAssert.class);
    }

    public static WorkbookAssert assertThat(Supplier<Workbook> supplier){
        return assertThat(supplier.get());
    }

    public static WorkbookAssert assertThat(Workbook workbook){
        return new WorkbookAssert(workbook);
    }

    public SheetAssert hasSheet(String sheetName) {
        isNotNull();
        Sheet sheet = actual.getSheet(sheetName);
        if(sheet == null){
            failWithMessage("Expecting workbook to contain sheet <%s> but doesn't.", sheetName);
        }
        return new SheetAssert(sheet);
    }

}
