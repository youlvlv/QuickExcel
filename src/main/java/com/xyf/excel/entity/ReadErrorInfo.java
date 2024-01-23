package com.xyf.excel.entity;

/**
 * 读取错误信息列表
 *
 * @author lizhiwei
 */
public class ReadErrorInfo {
	private int lineNumber;

	private String field;

	public int getLineNumber() {
		return lineNumber;
	}

	public void setLineNumber(int lineNumber) {
		this.lineNumber = lineNumber;
	}

	public String getField() {
		return field;
	}

	public void setField(String field) {
		this.field = field;
	}

	public ReadErrorInfo(int lineNumber, String field) {
		this.lineNumber = lineNumber;
		this.field = field;
	}

	public ReadErrorInfo() {
	}
}
