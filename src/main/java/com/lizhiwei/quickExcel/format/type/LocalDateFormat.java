package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;

public class LocalDateFormat implements ExcelFormatByType<LocalDate> {

	private static final DateTimeFormatter FMT1 = DateTimeFormatter.ofPattern("yyyy-MM-dd");


	private static final DateTimeFormatter FMT2 = DateTimeFormatter.ofPattern("yyyy/MM/dd");

	@Override
	public Class<LocalDate> getType() {
		return LocalDate.class;
	}

	@Override
	public String writer(LocalDate v) {
		return v.format(FMT1);
	}

	@Override
	public LocalDate ReadToExcel(String v) {
		if (v.length() >= 10){
			v = v.substring(0, 10);
		}
		DateTimeFormatter formatter;
		// 判断当前的解析方法
		if (v.contains("-")) {
			formatter = FMT1;
		} else if (v.contains("/")) {
			formatter = FMT2;
		} else {
			// 存在无法匹配的方案，直接返回 null
			return null;
		}
		try {
			return LocalDate.parse(v, formatter);
		} catch (DateTimeParseException e) {
			// 当前无法正常解析日期
			return null;
		}

	}
}
