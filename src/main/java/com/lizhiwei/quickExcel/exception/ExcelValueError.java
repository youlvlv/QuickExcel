package com.lizhiwei.quickExcel.exception;

import com.lizhiwei.quickExcel.entity.ReadErrorInfo;

import java.util.ArrayList;
import java.util.List;

/**
 * @author lizhiwei
 */
public class ExcelValueError extends RuntimeException {

    private List<ReadErrorInfo> errorInfos = new ArrayList<>();
    public ExcelValueError() {
    }

    public ExcelValueError(String message) {
        super(message);
    }

    public ExcelValueError(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelValueError(List<ReadErrorInfo> errorInfos){
        super("当前excel出现错误");
        this.errorInfos = errorInfos;
    }

    public ExcelValueError(Throwable cause) {
        super(cause);
    }

    public ExcelValueError(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

    public List<ReadErrorInfo> getErrorInfos() {
        return errorInfos;
    }
}
