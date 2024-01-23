package com.xyf.excel.entity;

import com.xyf.excel.format.ExcelFormatByType;
import com.xyf.excel.format.type.NullFormat;

import java.util.HashMap;

/**
 * 类型转换器存储
 */
public class ClassMap extends HashMap<Class<?>, ExcelFormatByType<?>> {


    @Override
    public ExcelFormatByType<?> get(Object key) {
        throw new RuntimeException("类型错误");
    }

    public ExcelFormatByType<?> get(Class<?> key) {
        if (super.containsKey(key)) {
            return superGet(key);
        } else if (super.containsKey(key.getSuperclass())) {
            return superGet(key.getSuperclass());
        } else {
            return new NullFormat();
        }
    }

    private ExcelFormatByType<?> superGet(Object key) {
        return super.get(key);
    }
}
