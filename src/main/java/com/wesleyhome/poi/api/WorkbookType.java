package com.wesleyhome.poi.api;

public enum WorkbookType {
    EXCEL_BIN("xls"),
    EXCEL_OPEN("xlsx"),
    EXCEL_STREAM("xlsx");


    private final String fileExtension;

    private WorkbookType(String fileExtension) {
        this.fileExtension = fileExtension;
    }

    public String getFileExtension() {
        return fileExtension;
    }
}
