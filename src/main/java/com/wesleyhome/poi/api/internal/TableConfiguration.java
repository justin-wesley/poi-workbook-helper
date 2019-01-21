package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.TableStyle;
import com.wesleyhome.poi.api.report.annotations.TotalsRowFunction;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;

import java.util.ArrayList;
import java.util.List;

public interface TableConfiguration {

    TableStyle getTableStyle();

    boolean hasTotalRow();

    default List<TotalsRowFunction> getTotalRowFunctions() {
        return new ArrayList();
    }

    String getTableName();
}
