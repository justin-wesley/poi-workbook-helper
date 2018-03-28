package com.wesleyhome.poi.api.internal;

import java.util.Map;
import java.util.TreeMap;

public class CellStyleManager {

    private Map<String, CellStyle> styleMap;

    CellStyleManager(){
        styleMap = new TreeMap<>();
    }

    CellStyle copy(String cellStyle) {
        return null;
    }
}
