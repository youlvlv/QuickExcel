package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

/**
 * String 转换器
 */
public class StringFormat implements ExcelFormatByType<String> {
    @Override
    public String writer(String v) {
        return v;
    }

    @Override
    public String ReadToExcel(String v) {
        return v;
    }

    @Override
    public Class<String> getType() {
        return String.class;
    }
}
