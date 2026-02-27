package com.lizhiwei.quickExcel.entity;

import java.io.InputStream;
import java.util.function.BiFunction;

public interface ImageFileSaveFunction extends BiFunction<InputStream, String, String> {
    @Override
    String apply(InputStream inputStream, String fileExtension);
}
