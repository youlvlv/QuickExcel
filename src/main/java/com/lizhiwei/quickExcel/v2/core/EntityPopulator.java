package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ParamType;
import com.lizhiwei.quickExcel.entity.PictureMap;
import com.lizhiwei.quickExcel.entity.Rule;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import org.apache.poi.ss.usermodel.Picture;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * 实体填充器
 * 负责将值填充到实体对象中
 */
public class EntityPopulator {
    
    /**
     * 填充实体对象
     * @param entity 实体对象
     * @param property 属性定义
     * @param value 值
     * @param pictureMap 图片映射（可选）
     * @param rowNum 行号
     * @throws ExcelValueException 值转换异常
     */
    public static void populate(Object entity, ExcelEntity property, Object value, 
                                 PictureMap pictureMap, int rowNum) throws ExcelValueException {
        try {
            // 执行验证规则
            if (value != null) {
                for (Rule rule : property.getRules()) {
                    rule.rule(value);
                }
            }
            
            // 根据参数类型设置值
            switch (property.getParamType()) {
                case FIELD -> setFieldValue(entity, property, value);
                case METHOD -> setMethodValue(entity, property, value);
                case IMAGE -> setImageValue(entity, property, pictureMap, rowNum);
            }
        } catch (NoSuchFieldException | IllegalAccessException | NoSuchMethodException | 
                 InvocationTargetException e) {
            throw new RuntimeException("设置字段值失败: " + property.getProperty(), e);
        }
    }
    
    /**
     * 设置字段值
     */
    private static void setFieldValue(Object entity, ExcelEntity property, Object value) 
            throws NoSuchFieldException, IllegalAccessException {
        Field field = entity.getClass().getDeclaredField(property.getProperty());
        field.setAccessible(true);
        field.set(entity, value);
    }
    
    /**
     * 设置方法值
     */
    private static void setMethodValue(Object entity, ExcelEntity property, Object value) 
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        String setterName = "set" + Pattern.compile("^.")
                .matcher(property.getProperty())
                .replaceFirst(m -> m.group().toUpperCase());
        Method method = entity.getClass().getMethod(setterName, property.getType());
        method.invoke(entity, value);
    }
    
    /**
     * 设置图片值
     */
    private static void setImageValue(Object entity, ExcelEntity property, 
                                       PictureMap pictureMap, int rowNum) 
            throws NoSuchFieldException, IllegalAccessException {
        if (pictureMap == null) {
            return;
        }
        
        Picture picture = pictureMap.get(new PictureMap.PictureKey(rowNum, property.getValue()));
        if (picture == null) {
            return;
        }
        
        Object imageValue = ImageExtractor.processImageField(pictureMap, rowNum, property);
        if (imageValue != null) {
            Field field = entity.getClass().getDeclaredField(property.getProperty());
            field.setAccessible(true);
            field.set(entity, EntityValueConverter.convert(imageValue.toString(), property, null));
        }
    }
}
