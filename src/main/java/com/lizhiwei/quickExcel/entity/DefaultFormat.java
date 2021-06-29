package com.lizhiwei.quickExcel.entity;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DefaultFormat implements ExcelFormat{


    @Override
    public  Object WriterToExcel(Object v) {
        String value = "";
        //判断属性的类型
        if (v instanceof String) {
            //String类型执行toString方法
            value = v.toString();
        } else if (v instanceof Date) {
            //时间类型，则转换时间
            DateFormat dFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm"); //HH表示24小时制；
            value = dFormat.format((Date) v);
            if (value.contains("08:00")) {
                value = value.substring(0, value.length() - 5);
            }
        } else if (v instanceof Number) {
            value = v.toString();
        }
        return value;
    }
}
