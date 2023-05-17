package com.lizhiwei.quickExcel.format.type;

import com.lizhiwei.quickExcel.format.ExcelFormatByType;
import com.lizhiwei.quickExcel.entity.Null;

/**
 * 未匹配中的转换器
 */
public class NullFormat implements ExcelFormatByType {
    @Override
    public String writer(Object v) {
        return v.toString();
    }

    @Override
    public Object ReadToExcel(String v) {
        return null;
    }

    @Override
    public Class getType() {
        return Null.class;
    }
}
