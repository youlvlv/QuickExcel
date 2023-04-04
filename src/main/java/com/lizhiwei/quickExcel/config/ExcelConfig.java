package com.lizhiwei.quickExcel.config;

import com.lizhiwei.quickExcel.entity.DefaultFormat;
import com.lizhiwei.quickExcel.entity.ExcelFormat;

import java.util.HashMap;
import java.util.Map;

public class ExcelConfig {
    private static final Map<Class<?>, ExcelFormat<?>> formatCache = new HashMap<>() {{
        put(DefaultFormat.class, new DefaultFormat());
    }};

    static public Map<Class<?>, ExcelFormat<?>> getFormatCache() {
        return formatCache;
    }

    /**
     * 新增默认的转换器 用于节省内存或存在特殊转换器（如：仅包含带参构造器
     * @param clazz 转换器类型
     * @param format 转换器实例
     * @param <T> 转换器
     */
    public <T> void addFormat(Class<T> clazz,ExcelFormat<T> format){
        formatCache.put(clazz,format);
    }
}
