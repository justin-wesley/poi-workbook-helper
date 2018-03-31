package com.wesleyhome.poi.api.assertions;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.assertj.core.api.AbstractAssert;
import org.assertj.core.api.Assert;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.function.Supplier;

import static java.util.Objects.requireNonNull;

public class WorkbookAssert extends AbstractAssert<WorkbookAssert, Workbook> {
    private WorkbookAssert(Workbook actual) {
        super(actual, WorkbookAssert.class);
    }

    public static WorkbookAssert assertThat(Supplier<Workbook> supplier){
        return assertThat(supplier.get());
    }

    public static WorkbookAssert assertThat(Workbook workbook){
        Path tempDir = null;
        Path workbookPath = null;
        try{
            tempDir = Files.createTempDirectory("junit5-workbook");
            workbookPath = tempDir.resolve("workbook.xlsx");
            try(OutputStream os = Files.newOutputStream(workbookPath)){
                workbook.write(os);
            }
            try (InputStream is = Files.newInputStream(workbookPath)) {
                Workbook readWorkbook = WorkbookFactory.create(is);
                return new WorkbookAssert(readWorkbook);
            }
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                Files.deleteIfExists(requireNonNull(workbookPath));
                Files.deleteIfExists(tempDir);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        }
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
