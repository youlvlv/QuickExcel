package com.lizhiwei.quickExcel.format;

import java.util.Map;

/**
 * excel读取导出拓展点
 *
 * @param <T>
 */
public interface ExcelFormat<T> extends ExcelFormatBase<T> {

	T ReadToExcel(String v);

	@Override
	default T ReadToExcel(String v, Map<String, String> objectMap){
		return ReadToExcel(v);
	}

	default void init() {

	}

}
