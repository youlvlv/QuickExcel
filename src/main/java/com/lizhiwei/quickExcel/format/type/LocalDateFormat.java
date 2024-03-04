package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

public class LocalDateFormat implements ExcelFormatByType<LocalDate> {

	private static final Map<String, DateTimeFormatter> FMT = new HashMap<>() {{
		put("-", DateTimeFormatter.ofPattern("yyyy-MM-dd"));
		put("/", DateTimeFormatter.ofPattern("yyyy/MM/dd"));
		put(".", DateTimeFormatter.ofPattern("yyyy.MM.dd"));
		put("年", DateTimeFormatter.ofPattern("yyyyn年MM月dd日"));
	}};

	@Override
	public Class<LocalDate> getType() {
		return LocalDate.class;
	}

	@Override
	public String writer(LocalDate v) {
		return v.format(FMT.get("-"));
	}

	@Override
	public LocalDate ReadToExcel(String v) {
		if (v.length() >= 10) {
			v = v.substring(0, 10);
		}
		AtomicReference<DateTimeFormatter> formatter = new AtomicReference<>();
		String finalV = v;
		// 判断当前的解析方法
		FMT.forEach((k, f) -> {
			if (finalV.contains(k)) {
				formatter.set(f);
			}
		});
		if (formatter.get() == null) {
			return null;
		}
		try {
			return LocalDate.parse(v, formatter.get());
		} catch (DateTimeParseException e) {
			// 当前无法正常解析日期
			return null;
		}

	}
}
