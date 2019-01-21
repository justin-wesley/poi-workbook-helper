package com.wesleyhome.poi.api.report.annotations;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;

import static org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction.*;

public enum TotalsRowFunction {
    F_NONE(NONE),
    F_SUM(SUM),
    F_MIN(MIN),
    F_MAX(MAX),
    F_AVERAGE(AVERAGE),
    F_COUNT(COUNT),
    F_STD_DEV(STD_DEV);

    private final STTotalsRowFunction.Enum functionEnum;

    TotalsRowFunction(STTotalsRowFunction.Enum functionEnum) {

        this.functionEnum = functionEnum;
    }

    public STTotalsRowFunction.Enum getFunctionEnum() {
        return functionEnum;
    }
}
