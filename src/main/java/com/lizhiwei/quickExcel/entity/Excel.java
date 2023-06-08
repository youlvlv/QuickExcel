package com.lizhiwei.quickExcel.entity;

import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormat;

import java.lang.annotation.*;

/**
 * 字段注解
 *
 * @author lizhiwei
 */
@Target({ElementType.FIELD,ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {
	/**
	 * 导入EXCEL匹配名称
	 *
	 * @return
	 */
	String value();

	/**
	 * 导出excel时显示的第二名称
	 *
	 * @return
	 */
	Class<? extends TopName> secondName() default DefaultTopName.class;

	/**
	 * 导出EXCEL时 匹配名称 非必填
	 *  已废弃
	 * @return
	 */
	@Deprecated
	String name() default "";

	/**
	 * 默认排序
	 *
	 * @return
	 */
	int index() default -1;

	/**
	 * 转换工具
	 *
	 * @return
	 */
	Class<? extends ExcelFormat> format() default DefaultFormat.class;

	/**
	 * 是否可读
	 * @return
	 */
	boolean isRead() default true;

	/**
	 * 是否可写
	 * @return
	 */
	boolean isWrite() default true;

	/**
	 * 当前字段导入方式
	 * @return
	 */
	ParamType type() default ParamType.FIELD;

	/**
	 * 导入时，是否允许非空，默认允许为空
	 * @return
	 */
	boolean isNotNull() default false;
}
