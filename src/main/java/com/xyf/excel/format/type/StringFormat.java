package com.xyf.excel.format.type;

import com.xyf.excel.format.ExcelFormatByType;

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
