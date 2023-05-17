package com.lizhiwei.quickExcel.format;

/**
 * 类型转换器接口
 * @author lizhiwei
 */
public interface ExcelFormatByType<T> extends ExcelFormat<T>{
    Class<T> getType();

    default String writerToExcel(Object v){
        return writer((T)v);
    }

    @Override
    default String WriterToExcel(T v){
        return writer(v);
    }

    String writer(T v);

    @Override
    T ReadToExcel(String v);
}
