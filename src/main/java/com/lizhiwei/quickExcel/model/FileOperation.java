package com.lizhiwei.quickExcel.model;

/**
 * 文件操作类
 */
@FunctionalInterface
public interface FileOperation {
    /**
     * 文件处理方法
     * @param model
     */
    void run(ExcelModel model);
}
