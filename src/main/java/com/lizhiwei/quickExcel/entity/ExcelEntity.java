package com.lizhiwei.quickExcel.entity;


import com.lizhiwei.quickExcel.format.ExcelFormat;

import java.lang.reflect.InvocationTargetException;

/**
 * 导出Excel实体类
 */
public class ExcelEntity {
	/**
	 * 名称
	 */
	private String title;
	/**
	 * 值
	 */
	private Integer value;
	/**
	 * 属性名
	 */
	private String property;
	/**
	 * 类型
	 */
	private Class<?> type;
	/**
	 * 转换器
	 */
	private ExcelFormat<?> format;
	/**
	 * 排序
	 */
	private int index;
	/**
	 * 顶部名称
	 */
	private Class<? extends TopName> topName = DefaultTopName.class;
	/**
	 * 字段类型
	 */
	private ParamType paramType = ParamType.FIELD;

	private boolean isRead = true;


	private boolean isWrite = true;


	private boolean isNotNull = false;

	public boolean isRead() {
		return isRead;
	}

	public void setRead(boolean read) {
		isRead = read;
	}

	public boolean isWrite() {
		return isWrite;
	}

	public void setWrite(boolean write) {
		isWrite = write;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public Integer getValue() {
		return value;
	}

	public void setValue(Integer value) {
		this.value = value;
	}

	public String getProperty() {
		return property;
	}

	public void setProperty(String property) {
		this.property = property;
	}

	public Class<?> getType() {
		return type;
	}

	public void setType(Class<?> type) {
		this.type = type;
	}

	public ExcelFormat<?> getFormat() {
		return format;
	}

	public void setFormat(ExcelFormat<?> format) {
		this.format = format;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public Class<? extends TopName> getTopName() {
		return topName;
	}

	public ParamType getParamType() {
		return paramType;
	}

	public void setParamType(ParamType paramType) {
		this.paramType = paramType;
	}

	public TopName getTopNameInt() {
		try {
			return topName.getDeclaredConstructor().newInstance();
		} catch (InstantiationException | NoSuchMethodException | InvocationTargetException |
		         IllegalAccessException e) {
			throw new RuntimeException(e);
		}
	}

	public boolean isNotNull() {
		return isNotNull;
	}

	public void setNotNull(boolean notNull) {
		isNotNull = notNull;
	}

	public ExcelEntity(Integer value, String title, ExcelFormat<?> format) {
		this.title = title;
		this.value = value;
		this.format = format;
	}

	public ExcelEntity(String value, String title, ExcelFormat<?> format, int index, Class<? extends TopName> topName, ParamType type) {
		this.title = title;
		this.property = value;
		this.format = format;
		this.index = index;
		this.topName = topName;
		this.paramType = type;
	}


	public ExcelEntity(String value, String title, ExcelFormat<?> format, int index, Class<? extends TopName> topName,Class clazz ,ParamType type, boolean isRead, boolean isWrite,boolean isNotNull) {
		this.title = title;
		this.property = value;
		this.format = format;
		this.index = index;
		this.topName = topName;
		this.paramType = type;
		this.isRead = isRead;
		this.isWrite = isWrite;
		this.type = clazz;
		this.isNotNull = isNotNull;
	}

	public ExcelEntity(ParamType index) {
		if (index == ParamType.INDEX) {
			this.title = "序号";
			this.property = "";
			this.paramType = index;
			this.type = Integer.class;
		}
	}

	public ExcelEntity(Integer value, String title) {
		this.title = title;
		this.value = value;
	}

	public ExcelEntity(String value, String title) {
		this.title = title;
		this.property = value;
	}

	public ExcelEntity() {
	}

}
