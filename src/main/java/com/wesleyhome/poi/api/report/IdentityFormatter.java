package com.wesleyhome.poi.api.report;

import java.util.function.Function;

public class IdentityFormatter implements Function<Object, Object> {

    @Override
    public Object apply(Object o) {
        return o;
    }
}
