package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;

public class BooleanFormat implements ExcelFormatByType<Boolean> {
    @Override
    public String writer(Boolean v) {
        if (v) {
            return "是";
        } else {
            return "否";
        }
    }

    @Override
    public Boolean ReadToExcel(String v) {
        if (v.contains("是")) {
            return true;
        } else if (v.contains("否")) {
            return false;
        }
        return Boolean.valueOf(v);
    }

    @Override
    public Class<Boolean> getType() {
        return Boolean.class;
    }
}
