package com.lizhiwei.quickExcel.model;

import java.io.OutputStream;
import java.util.function.Consumer;

/**
 * 文件操作类
 */
@FunctionalInterface
public interface FileOperation {
    /**
     * 文件处理方法
     * @param model
     */
    void run(Consumer<OutputStream> model);
}
