package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.report.annotations.ReportColumn;

import java.util.function.Function;

import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public interface ColumnConfiguration<T> extends Comparable<ColumnConfiguration<T>>{

    boolean isDisplayed();

    String getColumnHeader();

    Object getColumnValue(T value) ;

    String getColumn();

    ColumnType getColumnType();

    String getPropertyName();

    @Override
    default int compareTo(ColumnConfiguration<T> that){
        String thisColumnName = this.getColumn();
        String thatColumnName = that.getColumn();
        int nameCompare = thisColumnName.compareTo(thatColumnName);
        if(nameCompare == 0){
            return 0;
        }
        if(ReportColumn.NULL.equals(thisColumnName)){
            return 1;
        }
        if(ReportColumn.NULL.equals(thatColumnName)){
            return -1;
        }
        int thisColumnNum = convertColStringToIndex(thisColumnName);
        int thatColumnNum = convertColStringToIndex(thatColumnName);
        return Integer.compare(thisColumnNum, thatColumnNum);
    }
}
