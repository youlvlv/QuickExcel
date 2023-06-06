package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class LocalDateFormat implements ExcelFormatByType<LocalDate> {

    private static final DateTimeFormatter FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    @Override
    public Class<LocalDate> getType() {
        return LocalDate.class;
    }

    @Override
    public String writer(LocalDate v) {
        return v.format(FMT);
    }

    @Override
    public LocalDate ReadToExcel(String v) {
        v=v.substring(0,10);
        return LocalDate.parse(v,FMT);
    }
}
