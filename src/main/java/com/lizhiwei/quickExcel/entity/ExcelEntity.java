package com.lizhiwei.quickExcel.entity;


/**
 * 导出Excel实体类
 */
public class ExcelEntity {
    private String title;
    private String value;
    private String property;
    private String type;

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

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public ExcelEntity(String value, String title) {
        this.title = title;
        this.value = value;
    }

    public ExcelEntity() {
    }

}
