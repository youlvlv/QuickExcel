package com.lizhiwei.quickExcel.entity;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DefaultFormat implements ExcelFormat<Object> {


    @Override
    public String WriterToExcel(Object v) {
        if (v != null) {
            //判断属性的类型
            if (v instanceof String) {
                //String类型执行toString方法
                return v.toString();
            } else if (v instanceof Date) {
                //时间类型，则转换时间
                DateFormat dFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); //HH表示24小时制；
                return dFormat.format((Date) v);
            } else if (v instanceof Number) {
                return v.toString();
            } else if (v instanceof Boolean) {
                if ((boolean) v) {
                    return "是";
                } else {
                    return "否";
                }
            } else return v.toString();
        }
        return "";
    }

    @Override
    public Object ReadToExcel(String v) {
        return null;
    }

    public Object ReadToExcel(Class<?> type, String v) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        if (v == null|| v.equals("")) {
            return null;
        } else {
            if (type == String.class) {
                return v;
            } else if (type == Integer.class) {
                return Integer.valueOf(v);
            } else if (type == Long.class) {
                return Long.valueOf(v);
            } else if (type == Double.class) {
                return Double.valueOf(v);
            } else if (type == Boolean.class) {
                if (v.contains("是")) {
                    return true;
                } else if (v.contains("否")) {
                    return false;
                }
                return Boolean.valueOf(v);
            } else if (type == Date.class) {
                try {
                    return sdf.parse(v);
                } catch (ParseException e) {
                    throw new RuntimeException(e);
                }
            }
        }
        return null;
    }
}
