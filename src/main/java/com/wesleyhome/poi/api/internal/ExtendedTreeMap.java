package com.wesleyhome.poi.api.internal;

import java.util.Comparator;
import java.util.TreeMap;

public class ExtendedTreeMap<K,V> extends TreeMap<K,V> implements ExtendedMap<K,V> {

    public ExtendedTreeMap() {
    }

    public ExtendedTreeMap(Comparator<? super K> comparator) {
        super(comparator);
    }
}
