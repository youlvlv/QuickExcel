package com.xyf.excel.model;


import java.io.File;
/**
 * 上传文件
 */
@FunctionalInterface
public interface UploadFile {

    /**
     * 获取文件
     * @return
     */
    File getFile();
}
