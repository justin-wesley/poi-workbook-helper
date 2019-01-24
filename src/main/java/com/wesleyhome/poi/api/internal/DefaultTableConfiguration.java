package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.TableStyle;
import com.wesleyhome.poi.api.report.annotations.TotalsRowFunction;
import lombok.Builder;
import lombok.Data;
import lombok.Singular;

import java.util.List;

@Data
@Builder
public class DefaultTableConfiguration implements TableConfiguration {

    private TableStyle tableStyle;
    @Singular
    private List<TotalsRowFunction> totalsRowFunctions;

}
