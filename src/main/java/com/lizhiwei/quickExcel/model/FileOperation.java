package com.lizhiwei.quickExcel.model;

/**
 * 文件操作类
 */
@FunctionalInterface
public interface FileOperation {

    void run(ExcelModel model);
}
