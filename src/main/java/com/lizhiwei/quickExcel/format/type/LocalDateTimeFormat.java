package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class LocalDateTimeFormat implements ExcelFormatByType<LocalDateTime> {

    private static final DateTimeFormatter FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss ");

    @Override
    public Class<LocalDateTime> getType() {
        return LocalDateTime.class;
    }

    @Override
    public String writer(LocalDateTime v) {
        return v.format(FMT);
    }

    @Override
    public LocalDateTime ReadToExcel(String v) {
        return LocalDateTime.parse(v,FMT);
    }
}
