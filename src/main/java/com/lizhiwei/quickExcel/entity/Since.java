package com.lizhiwei.quickExcel.entity;

/**
 * 纵向数据合并
 */
public class Since {

    private final int startRow;
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

    public int getStartRow() {
        return startRow;
    }

    public Since(int startRow, int row, String title) {
        this.startRow = startRow;
        this.row = row;
        this.title = title;
    }
}
