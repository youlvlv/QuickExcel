package com.lizhiwei.quickExcel.exception;

public class IORunTimeException extends RuntimeException {
    public IORunTimeException() {
        super();
    }

    public IORunTimeException(String message) {
        super(message);
    }

    public IORunTimeException(String message, Throwable cause) {
        super(message, cause);
    }

    public IORunTimeException(Throwable cause) {
        super(cause);
    }
}
