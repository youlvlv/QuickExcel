package com.xyf.excel.format.type;

import com.xyf.excel.format.ExcelFormatByType;

public class DoubleFormat implements ExcelFormatByType<Double> {

    @Override
    public Class<Double> getType() {
        return Double.class;
    }

    @Override
    public String writer(Double v) {
        return v.toString();
    }

    @Override
    public Double ReadToExcel(String v) {
        return Double.valueOf(v);
    }
}
