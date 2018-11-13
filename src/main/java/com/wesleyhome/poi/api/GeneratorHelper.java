package com.wesleyhome.poi.api;


import java.util.Iterator;
import java.util.function.BiFunction;
import java.util.function.Function;

public class GeneratorHelper {

    public static <G, T> G iterate(G currentGenerator, Iterable<T> iterable, BiFunction<G, T, G> generatorFunction, Function<G, G> nextFunction) {
        G gen = currentGenerator;
        Iterator<T> iterator = iterable.iterator();
        while(iterator.hasNext()) {
            T next = iterator.next();
            gen = generatorFunction.apply(gen, next);
            if(iterator.hasNext()){
                gen = nextFunction.apply(gen);
            }
        }
        return gen;
    }
 }
