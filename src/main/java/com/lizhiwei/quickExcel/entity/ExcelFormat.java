package com.lizhiwei.quickExcel.entity;

/**
 * 旧版本兼容层，不推荐使用
 * 建议更换为
 * @see com.lizhiwei.quickExcel.format.ExcelFormat
 * @param <T>
 */
@Deprecated(forRemoval = true)
public interface ExcelFormat<T> extends com.lizhiwei.quickExcel.format.ExcelFormat<T> {
}
