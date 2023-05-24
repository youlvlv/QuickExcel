package com.lizhiwei.quickExcel.entity;

/**
 * 纵向数据合并
 */
public class Since {
    /**
     * 合并的行数
     */
    private final int row;
    /**
     * 属性名
     */
    private final String title;

    public int getRow() {
        return row;
    }

    public String getTitle() {
        return title;
    }

    public Since(int row, String title) {
        this.row = row;
        this.title = title;
    }
}
