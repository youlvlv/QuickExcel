package com.xyf.excel.format;

/**
 * excel读取导出拓展点
 *
 * @param <T>
 */
public interface ExcelFormat<T> extends ExcelFormatBase<T> {

	default void init() {

	}

}
