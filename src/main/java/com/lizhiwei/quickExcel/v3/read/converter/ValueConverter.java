package com.lizhiwei.quickExcel.v3.read.converter;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ParamType;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.regex.Pattern;

/**
 * 值转换器
 * 负责将字符串值转换为实体类的属性类型
 */
public class ValueConverter {
    
    private static final Pattern SET_METHOD_PATTERN = Pattern.compile("^.");
    
    /**
     * 转换值并设置到实体对象
     */
    public static void convertAndSet(Object obj, ExcelEntity property, String value) throws Exception {
        if (value == null || value.trim().isEmpty()) {
            if (property.isNotNull()) {
                throw new ExcelValueException(property.getTitle() + "不能为空");
            }
            return;
        }
        
        Object convertedValue = convert(property, value);
        setFieldValue(obj, property, convertedValue);
    }
    
    /**
     * 转换字符串值到目标类型
     */
    public static Object convert(ExcelEntity property, String value) {
        if (value == null || value.trim().isEmpty()) {
            return null;
        }
        
        ExcelFormatBase<?> format = property.getFormat();
        if (format != null) {
            try {
                if (format instanceof DefaultFormat) {
                    return ((DefaultFormat) format).ReadToExcel(property.getType(), value);
                }
                return format.ReadToExcel(value, null);
            } catch (Exception e) {
                throw new ExcelValueException(property.getTitle() + "格式错误: " + value, e);
            }
        }
        
        return convertByType(property.getType(), value, property.getTitle());
    }
    
    /**
     * 根据类型进行转换
     */
    private static Object convertByType(Class<?> type, String value, String title) {
        try {
            if (type == String.class) {
                return value;
            } else if (type == Integer.class || type == int.class) {
                return Integer.parseInt(value);
            } else if (type == Long.class || type == long.class) {
                return Long.parseLong(value);
            } else if (type == Double.class || type == double.class) {
                return Double.parseDouble(value);
            } else if (type == BigDecimal.class) {
                return new BigDecimal(value);
            } else if (type == Boolean.class || type == boolean.class) {
                return Boolean.parseBoolean(value);
            } else if (type == java.util.Date.class) {
                return new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(value);
            } else if (type == java.time.LocalDate.class) {
                return java.time.LocalDate.parse(value);
            } else if (type == java.time.LocalDateTime.class) {
                return java.time.LocalDateTime.parse(value.replace(" ", "T"));
            } else if (type == java.time.LocalTime.class) {
                return java.time.LocalTime.parse(value);
            } else if (type == byte[].class) {
                // 支持 byte[] 类型的字段存储图片数据
                return value.getBytes();
            }
        } catch (Exception e) {
            throw new ExcelValueException(title + "转换失败: " + value, e);
        }
        
        return value;
    }
    
    /**
     * 设置字段值（支持字段和Setter方法）
     */
    public static void setFieldValue(Object obj, ExcelEntity property, Object value) throws Exception {
        if (value == null) {
            return;
        }
        
        switch (property.getParamType()) {
            case FIELD:
                setField(obj, property.getProperty(), value);
                break;
            case METHOD:
                setMethod(obj, property.getProperty(), property.getType(), value);
                break;
            case IMAGE:
                // 图片类型字段的处理在更高层完成（如 SaxExcelReader 和 DomExcelReader）
                // 这里只处理图片名称或标识的字符串值
                if (value instanceof String) {
                    setField(obj, property.getProperty(), value);
                }
                break;
            default:
                // 默认按字段处理
                setField(obj, property.getProperty(), value);
                break;
        }
    }
    
    /**
     * 设置图片字段值
     * 
     * @param obj 实体对象
     * @param property 实体属性
     * @param imageData 图片数据
     * @throws Exception 反射操作异常
     */
    public static void setImageFieldValue(Object obj, ExcelEntity property, ImageData imageData) throws Exception {
        if (property.getParamType() != ParamType.IMAGE) {
            return;
        }
        
        Class<?> fieldType = property.getType();
        Object valueToSet = null;
        
        // 根据字段类型设置不同的值
        if (fieldType == byte[].class) {
            // byte[] 类型：设置图片字节数组
            valueToSet = imageData.getData();
        } else if (fieldType == ImageData.class) {
            // ImageData 类型：直接设置 ImageData 对象
            valueToSet = imageData;
        } else if (fieldType == String.class) {
            // String 类型：设置图片名称或路径
            valueToSet = imageData.getImageName() != null ? imageData.getImageName() : imageData.getFileName();
        } else if (fieldType == java.io.InputStream.class) {
            // InputStream 类型：设置图片数据流
            valueToSet = imageData.getData();
        }
        
        if (valueToSet != null) {
            try {
                setField(obj, property.getProperty(), valueToSet);
            } catch (NoSuchFieldException e) {
                // 如果字段不存在，尝试使用 setter 方法
                setMethod(obj, property.getProperty(), fieldType, valueToSet);
            }
        }
    }
    
    /**
     * 通过反射设置字段值
     */
    private static void setField(Object obj, String fieldName, Object value) throws Exception {
        Field field = obj.getClass().getDeclaredField(fieldName);
        field.setAccessible(true);
        field.set(obj, value);
    }
    
    /**
     * 通过反射调用Setter方法
     */
    private static void setMethod(Object obj, String propertyName, Class<?> paramType, Object value) throws Exception {
        String setMethodName = "set" + SET_METHOD_PATTERN.matcher(propertyName)
                .replaceFirst(m -> m.group().toUpperCase());
        Method method = obj.getClass().getMethod(setMethodName, paramType);
        method.invoke(obj, value);
    }
    
    /**
     * 创建实体类新实例
     */
    public static <T> T newInstance(Class<T> entityClass) throws Exception {
        return entityClass.getDeclaredConstructor().newInstance();
    }
}
