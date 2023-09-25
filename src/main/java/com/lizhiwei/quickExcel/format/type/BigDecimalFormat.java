package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.math.BigDecimal;

public class BigDecimalFormat implements ExcelFormatByType<BigDecimal> {

    @Override
    public Class<BigDecimal> getType() {
        return BigDecimal.class;
    }

    @Override
    public String writer(BigDecimal v) {
        return v.toString();
    }

    @Override
    public BigDecimal ReadToExcel(String v) {
        return new BigDecimal(v);
    }
}
