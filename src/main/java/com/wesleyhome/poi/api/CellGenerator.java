package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Arrays;
import java.util.Iterator;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

public interface CellGenerator extends WorkbookCreator{
    CellGenerator havingValue(Object cellValue);

    boolean hasValue();

    CellGenerator sumOfPreviousXRows(int numberOfRows);

    CellGenerator withFormula(String formula, Object... args);

    CellGenerator usingStyle(String cellStyle);

    CellGenerator withDateFormat();

    CellGenerator withIntegerFormat();

    CellGenerator withCurrencyFormat();

    CellGenerator withNumericStyle();

    CellGenerator withAccountingFormat();

    CellGenerator withBackgroundColor(IndexedColors color);

    CellGenerator withFontColor(IndexedColors color);

    CellGenerator hide();

    CellGenerator isBold();

    CellGenerator notBold();

    CellGenerator as(String cellStyleName);

    CellGenerator withNoBackgroundColor();

    CellGenerator cellWidth(float width);

    CellGenerator withAllBorders(BorderStyle borderStyle, IndexedColors color);

    CellGenerator withTopBorder(BorderStyle borderStyle, IndexedColors color);

    CellGenerator withBottomBorder(BorderStyle borderStyle, IndexedColors color);

    CellGenerator withLeftBorder(BorderStyle borderStyle, IndexedColors color);

    CellGenerator withRightBorder(BorderStyle borderStyle, IndexedColors color);

    CellGenerator mergeWithNextXCells(int numberOfCellsToMerge);

    CellGenerator withHorizontalAlignment(HorizontalAlignment horizontalAlignment);

    CellGenerator withWrappedText();

    CellGenerator withoutWrappedText();

    CellGenerator autosize();

    CellAddress cellAddress();

    default CellGenerator applyIf(Predicate<CellGenerator> filter, Consumer<CellGenerator> consumer) {
        if(filter.test(this)){
            consumer.accept(this);
        }
        return this;
    }

    default CellGenerator generateCell(Function<CellGenerator, CellGenerator> cellGeneratorFunction) {
        return cellGeneratorFunction.apply(this);
    }

    default <T1> CellGenerator generateCells(Iterable<T1> iterable, BiFunction<CellGenerator, T1, CellGenerator> cellGeneratorFunction) {
        return GeneratorHelper.iterate(this, iterable, cellGeneratorFunction, WorkbookCreator::nextCell);
    }

    default <T1> CellGenerator generateCells(T1[] array, BiFunction<CellGenerator, T1, CellGenerator> cellGeneratorFunction) {
        return generateCells(Arrays.asList(array), cellGeneratorFunction);
    }

    CellRangeAddress rangeAddress(int endRow, int columnNum);
}
