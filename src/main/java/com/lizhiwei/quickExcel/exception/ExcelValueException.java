package com.lizhiwei.quickExcel.exception;

import com.lizhiwei.quickExcel.entity.ReadErrorInfo;

import java.util.ArrayList;
import java.util.List;

/**
 * @author lizhiwei
 */
public class ExcelValueException extends ExcelBaseException {


	public ExcelValueException(String message) {
		super(message);
	}

	public ExcelValueException(String message, Throwable cause) {
		super(message, cause);
	}


	public ExcelValueException(Throwable cause) {
		super(cause);
	}

	@Override
	public List<ReadErrorInfo> getErrorInfos() {
		return new ArrayList<>();
	}


}
