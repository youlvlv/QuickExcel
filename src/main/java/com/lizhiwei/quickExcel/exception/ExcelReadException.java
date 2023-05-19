package com.lizhiwei.quickExcel.exception;

import com.lizhiwei.quickExcel.entity.ReadErrorInfo;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * 读取excel错误
 * @author lizhiwei
 */
public class ExcelReadException extends ExcelBaseException {
	private List<ReadErrorInfo> errorInfos = new ArrayList<>();
	public ExcelReadException() {
	}

	public ExcelReadException(String message) {
		super(message);
	}

	public ExcelReadException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelReadException(List<ReadErrorInfo> errorInfos){
		super("当前excel出现错误");
		this.errorInfos = errorInfos;
	}

	public ExcelReadException(Throwable cause) {
		super(cause);
	}

	public ExcelReadException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public List<ReadErrorInfo> getErrorInfos() {
		return errorInfos;
	}
}
