package com.xyf.excel.entity;

public @interface ExcelGroup {


	Class paramType();


	/**
	 * 当前字段导入方式
	 * @return
	 */
	ParamType type() default ParamType.METHOD;


}
