package com.kay.xlsexporter;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

import java.util.function.Function;

@RequiredArgsConstructor
@Getter
public class Column<T> {
    private final String header;
    private final Function<T, ?> valueProvider;
    private final Integer width;
}
