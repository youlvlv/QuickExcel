package com.xyf.excel.format.type;

import com.xyf.excel.format.ExcelFormatByType;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateFormat implements ExcelFormatByType<Date> {

    private static final java.text.DateFormat dFormat = new SimpleDateFormat("yyyy-MM-dd"); //HH表示24小时制

    @Override
    public String writer(Date v) {
        return dFormat.format(v);
    }

    @Override
    public Date ReadToExcel(String v) {
        try {
            return dFormat.parse(v);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public Class<Date> getType() {
        return Date.class;
    }
}
