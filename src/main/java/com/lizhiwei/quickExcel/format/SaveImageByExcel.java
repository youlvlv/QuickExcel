package com.lizhiwei.quickExcel.format;

import java.io.InputStream;
import java.util.function.BiFunction;

public interface SaveImageByExcel extends BiFunction<InputStream, String, String> {

	@Override
	String apply(InputStream inputStream, String type);
}
