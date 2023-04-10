package com.lizhiwei.quickExcel.config;

import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatByType;

import java.util.HashMap;
import java.util.Map;

public class ExcelConfig {
    private static final HashMap<Class<?>, ExcelFormat<?>> formatCache = new HashMap<>() {{
        put(DefaultFormat.class, new DefaultFormat());
    }};

    static public HashMap<Class<?>, ExcelFormat<?>> getFormatCache() {
        return (HashMap<Class<?>, ExcelFormat<?>>) formatCache.clone();
    }

    /**
     * 新增默认的转换器 用于节省内存或存在特殊转换器（如：仅包含带参构造器
     * @param clazz 转换器类型
     * @param format 转换器实例
     * @param <T> 转换器
     */
    public static <T> void addFormat(Class<T> clazz,ExcelFormat<T> format){
        formatCache.put(clazz,format);
    }

    /**
     * 移除转换器
     * @param clazz
     */
    public static void removeFormat(Class<?> clazz){
        formatCache.remove(clazz);
    }


    public static <T> void addTypeFormat(Class<T> clazz, ExcelFormatByType<T> format) {
        DefaultFormat.map.put(clazz,format);
    }

    public static void removeTypeFormat(Class<?> clazz) {
        DefaultFormat.map.remove(clazz);
    }
}
