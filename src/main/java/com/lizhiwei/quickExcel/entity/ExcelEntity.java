package com.lizhiwei.quickExcel.entity;

import com.chinatechstar.component.commons.utils.PageData;

/**
 * 导出Excel实体类
 */
public class ExcelEntity {
    private String title;
    private String value;
    private String property;
    private Class<?> type;
    private ExcelFormat<?> format;
    private int index;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
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

    public ExcelEntity(String value, String title, ExcelFormat<?> format) {
        this.title = title;
        this.value = value;
        this.format = format;
    }

    public ExcelEntity(String value, String title, ExcelFormat<?> format,int index) {
        this.title = title;
        this.value = value;
        this.format = format;
        this.index = index;
    }

    public ExcelEntity(String value, String title) {
        this.title = title;
        this.value = value;
    }

    public ExcelEntity() {
    }

    public PageData toPageData() {
        PageData a = new PageData();
        a.put("value", this.value);
        a.put("title", this.title);
        a.put("property",this.property);
        return a;
    }
}
