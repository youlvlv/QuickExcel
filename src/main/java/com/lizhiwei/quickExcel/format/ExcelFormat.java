package com.lizhiwei.quickExcel.format;

/**
 * excel读取导出拓展点
 *
 * @param <T>
 */
public interface ExcelFormat<T> extends Cloneable {
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
	 * @return 属性
	 */
	T ReadToExcel(String v);

	/**
	 * 初始化构造器，每次构建excel单元格转换器时都会重新调用，方便同步数据
	 * 该方法仅会在ExcelConfig中添加的转换器被初始化时调用
	 */
	default ExcelFormat<T> init() {
		return this;
	}

	default void over() {

	}
}
