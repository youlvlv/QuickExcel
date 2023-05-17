package com.lizhiwei.quickExcel.entity;

/**
 * 纵向数据合并
 */
public class Since {

    private final int row;

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
