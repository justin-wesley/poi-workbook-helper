package com.wesleyhome.poi.api.report;

import org.apache.commons.lang3.concurrent.ConcurrentException;
import org.apache.commons.lang3.concurrent.LazyInitializer;

public abstract class SafeLazyInitializer<T> extends LazyInitializer<T> {
    @Override
    protected final T initialize() throws ConcurrentException {
        try {
            return safeInitialize();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    protected abstract T safeInitialize();

    @Override
    public T get() {
        try {
            return super.get();
        } catch (ConcurrentException e) {
            throw new RuntimeException(e);
        }
    }
}
