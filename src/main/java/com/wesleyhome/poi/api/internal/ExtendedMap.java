package com.wesleyhome.poi.api.internal;

import java.util.NavigableMap;
import java.util.function.Supplier;

public interface ExtendedMap<K,V> extends NavigableMap<K,V> {

    default V getOrDefault(K key, Supplier<V> supplier){
        V value = get(key);
        if(value == null){
            value = supplier.get();
            put(key, value);
        }
        return value;
    }
}
