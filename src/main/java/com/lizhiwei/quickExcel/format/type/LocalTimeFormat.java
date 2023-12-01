package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

public class LocalTimeFormat implements ExcelFormatByType<LocalTime> {

	private static final DateTimeFormatter FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

	@Override
	public Class<LocalTime> getType() {
		return LocalTime.class;
	}

	@Override
	public String writer(LocalTime v) {
		return v.format(FMT);
	}

	@Override
	public LocalTime ReadToExcel(String v) {
		return LocalTime.parse(v, FMT);
	}
}
