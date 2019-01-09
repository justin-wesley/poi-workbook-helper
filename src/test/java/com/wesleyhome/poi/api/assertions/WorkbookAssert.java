package com.wesleyhome.poi.api.assertions;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.assertj.core.api.AbstractAssert;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.function.Supplier;

import static java.util.Objects.requireNonNull;

public class WorkbookAssert extends AbstractAssert<WorkbookAssert, Workbook> {
    private WorkbookAssert(Workbook actual) {
        super(actual, WorkbookAssert.class);
    }

    public static WorkbookAssert assertThat(Supplier<Workbook> supplier) {
        return assertThat(supplier.get());
    }

    public static WorkbookAssert assertThat(Workbook workbook) {
        Path workbookPath = null;
        try {
            workbookPath = Files.createTempFile("junit5", ".xlsx");
            System.out.println(workbookPath.toAbsolutePath());
            try (OutputStream os = Files.newOutputStream(workbookPath)) {
                workbook.write(os);
            }
            try (InputStream is = Files.newInputStream(workbookPath)) {
                Workbook readWorkbook = WorkbookFactory.create(is);
                return new WorkbookAssert(readWorkbook);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(workbookPath.toFile());
                } else {
                    Files.deleteIfExists(requireNonNull(workbookPath));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public SheetAssert hasSheet(String sheetName) {
        isNotNull();
        Sheet sheet = actual.getSheet(sheetName);
        if (sheet == null) {
            failWithMessage("Expecting workbook to contain sheet <%s> but doesn't.", sheetName);
        }
        return new SheetAssert(sheet);
    }

}
