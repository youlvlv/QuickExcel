package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;

import java.util.Map;

/**
 * 实体值转换器
 * 负责将字符串值转换为实体字段的类型
 */
public class EntityValueConverter {
    
    /**
     * 转换值为实体字段类型
     * @param value 字符串值
     * @param property Excel 实体属性
     * @param objectMap 整行数据（用于多字段格式化）
     * @return 转换后的值
     */
    public static Object convert(String value, ExcelEntity property, Map<String, String> objectMap) {
        if (value == null || value.isEmpty()) {
            return null;
        }
        
        Class<?> type = property.getType();
        ExcelFormatBase<?> format = property.getFormat();
        
        try {
            if (format instanceof DefaultFormat) {
                return ((DefaultFormat) format).ReadToExcel(type, value);
            }
            return format.ReadToExcel(value, objectMap);
        } catch (Exception e) {
            throw new ExcelValueException(property.getTitle() + "错误", e);
        }
    }
    
    /**
     * 格式化值（仅格式化，不转换类型）
     * @param property Excel 实体属性
     * @param cellValue 单元格值
     * @return 格式化后的字符串值
     */
    public static String formatValue(ExcelEntity property, String cellValue) {
        if (property.getFormat() == null) {
            return cellValue;
        }
        
        ExcelFormatBase<?> format = property.getFormat();
        try {
            if (format instanceof DefaultFormat) {
                return ((DefaultFormat) format).ReadToExcel(String.class, cellValue).toString();
            }
            return format.ReadToExcel(cellValue, null).toString();
        } catch (Exception e) {
            throw new ExcelValueException(property.getTitle() + "错误", e);
        }
    }
}
