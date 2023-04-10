package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class LocalDateFormat implements ExcelFormatByType<LocalDate> {

    public static final DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    @Override
    public Class<LocalDate> getType() {
        return LocalDate.class;
    }

    @Override
    public String writer(LocalDate v) {
        return v.format(fmt);
    }

    @Override
    public LocalDate ReadToExcel(String v) {
        return LocalDate.parse(v,fmt);
    }
}
