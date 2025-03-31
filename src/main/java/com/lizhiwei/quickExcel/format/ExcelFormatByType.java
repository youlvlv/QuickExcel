package com.lizhiwei.quickExcel.format;

import java.util.Map;

/**
 * 类型转换器接口
 * @author lizhiwei
 */
public interface ExcelFormatByType<T> extends ExcelFormatBase<T>{
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
    default T ReadToExcel(String v, Map<String, String> objectMap){
        return ReadToExcel(v);
    }

    T ReadToExcel(String v);
}
