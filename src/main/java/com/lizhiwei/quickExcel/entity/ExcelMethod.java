package com.lizhiwei.quickExcel.entity;

import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormat;

import java.lang.annotation.*;

/**
 * 方法注解
 *
 * @author lizhiwei
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelMethod {
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
	 *
	 * @return
	 */
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
}
