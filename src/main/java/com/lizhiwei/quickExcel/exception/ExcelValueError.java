package com.lizhiwei.quickExcel.exception;

import com.lizhiwei.quickExcel.entity.ReadErrorInfo;

import java.util.ArrayList;
import java.util.List;

/**
 * @author lizhiwei
 */
public class ExcelValueError extends ExcelBaseException {


	public ExcelValueError(String message) {
		super(message);
	}

	public ExcelValueError(String message, Throwable cause) {
		super(message, cause);
	}


	public ExcelValueError(Throwable cause) {
		super(cause);
	}

	@Override
	public List<ReadErrorInfo> getErrorInfos() {
		return new ArrayList<>();
	}


}
