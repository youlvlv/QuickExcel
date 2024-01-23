package com.xyf.excel.entity;

/**
 * 旧版本兼容层，不推荐使用
 * 建议更换为
 * @see com.xyf.excel.format.ExcelFormat
 * @param <T>
 */
@Deprecated(forRemoval = true)
public interface ExcelFormat<T> extends com.xyf.excel.format.ExcelFormat<T> {
}
