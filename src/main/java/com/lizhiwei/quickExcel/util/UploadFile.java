package com.lizhiwei.quickExcel.util;


import java.io.File;

/**
 * 上传文件
 */
public abstract class UploadFile {

    private File file;


    protected UploadFile(File file) {
        this.file = file;
    }

    public File getFile() {
        return file;
    }
}
