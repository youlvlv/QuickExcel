package com.lizhiwei.quickExcel.entity;

public interface ExcelFormatByType<T> extends ExcelFormat<T>{
    Class<T> getType();
}
