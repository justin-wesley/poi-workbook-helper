package com.wesleyhome.poi.api.internal;

import lombok.EqualsAndHashCode;
import lombok.ToString;

import java.lang.reflect.Proxy;

@EqualsAndHashCode
@ToString
class Immutable<T> {
    private final T value;

    static <S> S immutable(S value){
        return value;
    }

    private Immutable(T value) {
        this.value = value;
    }

    private T get(){
        return value;
    }
}
