package com.lizhiwei.quickExcel.exception;

import com.lizhiwei.quickExcel.entity.ReadErrorInfo;

import java.util.List;

public abstract class ExcelBaseException extends RuntimeException{

	public ExcelBaseException() {
		super();
	}

	public ExcelBaseException(String message) {
		super(message);
	}

	public ExcelBaseException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelBaseException(Throwable cause) {
		super(cause);
	}

	protected ExcelBaseException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}


	public abstract List<ReadErrorInfo> getErrorInfos();
}
