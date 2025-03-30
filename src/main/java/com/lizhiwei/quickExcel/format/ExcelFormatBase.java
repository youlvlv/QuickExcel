package com.lizhiwei.quickExcel.format;

import java.util.Map;

public interface ExcelFormatBase<T> extends Cloneable {
	/**
	 * 读取实体类属性至excel值
	 *
	 * @param v
	 * @return
	 */
	String WriterToExcel(T v);

	/**
	 * 读取excel的值转换至实体类属性类型
	 *
	 * @param v 值
	 * @param oobjectMap 属性名与属性值
	 * @return 属性
	 */
	T ReadToExcel(String v, Map<String,String> objectMap);



	/**
	 * 初始化构造器，每次构建excel单元格转换器时都会重新调用，方便同步数据
	 * 该方法仅会在ExcelConfig中添加的转换器被初始化时调用
	 * 高级别的初始化方法，会调用 init 方法
	 */
	default ExcelFormatBase<T> initFormat() {
		this.init();
		return this;
	}

	/**
	 * 初始化构造器，每次构建excel单元格转换器时都会重新调用，方便同步数据
	 * 该方法仅会在ExcelConfig中添加的转换器被初始化时调用
	 * 低级别的初始化方法
	 */
	default void init() {
	}



	default void over() {

	}
}
