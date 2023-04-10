package com.lizhiwei.quickExcel.format;

public interface ExcelFormatByType<T> extends ExcelFormat<T>{
    Class<T> getType();

    @Override
    default String WriterToExcel(Object v){
        return writer((T)v);
    }

    String writer(T v);

    @Override
    T ReadToExcel(String v);
}
