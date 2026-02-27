package com.lizhiwei.quickExcel.v2.core;

import java.util.Map;

/**
 * 空行检查器
 * 负责检查行是否为空行
 */
public class EmptyRowChecker {
    
    private static final int DEFAULT_EMPTY_THRESHOLD = 3;
    
    private int emptySize = 0;
    private final int threshold;
    
    public EmptyRowChecker() {
        this.threshold = DEFAULT_EMPTY_THRESHOLD;
    }
    
    public EmptyRowChecker(int threshold) {
        this.threshold = threshold;
    }
    
    /**
     * 检查当前行是否为空行
     * @param nonEmptyFieldCount 非空字段数量
     * @return true 表示达到空行阈值，应停止读取
     */
    public boolean checkEmptyRow(int nonEmptyFieldCount) {
        if (nonEmptyFieldCount == 0) {
            emptySize++;
            return emptySize > threshold;
        } else {
            emptySize = 0;
            return false;
        }
    }
    
    /**
     * 重置计数器
     */
    public void reset() {
        emptySize = 0;
    }
    
    /**
     * 获取当前连续空行数
     */
    public int getEmptySize() {
        return emptySize;
    }
}
