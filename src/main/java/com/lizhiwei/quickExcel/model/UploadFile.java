package com.lizhiwei.quickExcel.model;


import java.io.File;

/**
 * 上传文件
 */
public abstract class UploadFile {

    private File file;


    protected UploadFile(File file) {
        this.file = file;
    }

    /**
     * 获取文件
     * @return
     */
    public File getFile() {
        return file;
    }
}
