package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

public class IntegerFormat implements ExcelFormatByType<Integer> {

    @Override
    public Class<Integer> getType() {
        return Integer.class;
    }

    @Override
    public String writer(Integer v) {
        return v.toString();
    }

    @Override
    public Integer ReadToExcel(String v) {
        return Integer.valueOf(v);
    }
}
