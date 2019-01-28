package com.wesleyhome.poi.api.report;

import org.apache.commons.lang3.concurrent.ConcurrentException;
import org.apache.commons.lang3.concurrent.LazyInitializer;

import java.util.function.Supplier;

public abstract class SupplierInitializer<V, T> extends LazyInitializer<T> {

    private final Supplier<V> valueSupplier;

    public SupplierInitializer(Supplier<V> valueSupplier) {
        this.valueSupplier = valueSupplier;
    }

    @Override
    protected final T initialize() throws ConcurrentException {
        try {
            return safeInitialize(this.valueSupplier.get());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    protected abstract T safeInitialize(V v);

    @Override
    public T get() {
        try {
            return super.get();
        } catch (ConcurrentException e) {
            throw new RuntimeException(e);
        }
    }
}
