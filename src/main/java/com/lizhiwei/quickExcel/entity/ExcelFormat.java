package com.lizhiwei.quickExcel.entity;

/**
 * excel读取导出拓展点
 * @param <T>
 */
public interface ExcelFormat<T> {
    /**
     * 读取实体类属性至excel值
     *
     * @param v
     * @return
     */
    String WriterToExcel(T v);

    /**
     * 读取excel的值转换至实体类属性类型
     *
     * @param v 值
     * @return 属性
     */
    T ReadToExcel(String v);
}
