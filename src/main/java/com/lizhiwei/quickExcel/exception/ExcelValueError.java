package com.lizhiwei.quickExcel.exception;

public class ExcelValueError extends RuntimeException {
    public ExcelValueError() {
    }

    public ExcelValueError(String message) {
        super(message);
    }

    public ExcelValueError(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelValueError(Throwable cause) {
        super(cause);
    }

    public ExcelValueError(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
