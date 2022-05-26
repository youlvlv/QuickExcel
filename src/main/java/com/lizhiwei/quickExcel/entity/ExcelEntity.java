package com.lizhiwei.quickExcel.entity;


import java.lang.reflect.InvocationTargetException;

/**
 * 导出Excel实体类
 */
public class ExcelEntity {
    /**
     * 名称
     */
    private String title;
    /**
     * 值
     */
    private Integer value;
    /**
     * 属性名
     */
    private String property;
    /**
     * 类型
     */
    private Class<?> type;
    /**
     * 转换器
     */
    private ExcelFormat<?> format;
    /**
     * 排序
     */
    private int index;
    /**
     * 顶部名称
     */
    private Class<? extends TopName> topName;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Integer getValue() {
        return value;
    }

    public void setValue(Integer value) {
        this.value = value;
    }

    public String getProperty() {
        return property;
    }

    public void setProperty(String property) {
        this.property = property;
    }

    public Class<?> getType() {
        return type;
    }

    public void setType(Class<?> type) {
        this.type = type;
    }

    public ExcelFormat<?> getFormat() {
        return format;
    }

    public void setFormat(ExcelFormat<?> format) {
        this.format = format;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public Class<? extends TopName> getTopName() {
        return topName;
    }

    public TopName getTopNameInt() {
        try {
            return topName.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | NoSuchMethodException | InvocationTargetException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    public ExcelEntity(Integer value, String title, ExcelFormat<?> format) {
        this.title = title;
        this.value = value;
        this.format = format;
    }

    public ExcelEntity(String value, String title, ExcelFormat<?> format, int index, Class<? extends TopName> topName) {
        this.title = title;
        this.property = value;
        this.format = format;
        this.index = index;
        this.topName = topName;
    }

    public ExcelEntity(Integer value, String title) {
        this.title = title;
        this.value = value;
    }

    public ExcelEntity(String value, String title) {
        this.title = title;
        this.property = value;
    }

    public ExcelEntity() {
    }

}
