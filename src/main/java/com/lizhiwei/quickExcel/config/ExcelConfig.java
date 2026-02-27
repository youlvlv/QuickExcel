package com.lizhiwei.quickExcel.config;

import com.lizhiwei.quickExcel.entity.ImageFileSaveFunction;
import com.lizhiwei.quickExcel.format.*;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Optional;
import java.util.function.BiFunction;

/**
 * 全局配置项
 */
public class ExcelConfig {

	private static final HashMap<Class<?>, ExcelFormatBase<?>> formatCache = new HashMap<>() {{
		put(DefaultFormat.class, new DefaultFormat());
	}};
	private static ImageFileSaveFunction imageFileFunction;
	
	/**
	 * Excel 读取引擎版本
	 */
	public enum ReadEngine {
		/**
		 * V2 引擎 - 基于 Apache POI（默认）
		 */
		V2,
		/**
		 * V3 引擎 - 基于 SAX/DOM 解析，内存占用更低
		 */
		V3
	}
	
	private static ReadEngine defaultReadEngine = ReadEngine.V2;

	static public HashMap<Class<?>, ExcelFormatBase<?>> getFormatCache() {
		return new HashMap<>(formatCache);
	}

    /**
     * 新增默认的转换器 用于节省内存或存在特殊转换器（如：仅包含带参构造器
     * @param format 转换器实例
     * @param <T> 转换器
     */
    public static <T> void addFormat(ExcelFormat<T> format){
        formatCache.put(format.getClass(),format);
    }

    public static <T> void addFormat(ExcelMoreFormat<T> format){
        formatCache.put(format.getClass(),format);
    }

	/**
	 * 移除转换器
	 *
	 * @param clazz
	 */
	public static void removeFormat(Class<?> clazz) {
		formatCache.remove(clazz);
	}


	public static <T> void addTypeFormat(Class<T> clazz, ExcelFormatByType<T> format) {
		DefaultFormat.CLASS_FORMAT_MAP.put(clazz, format);
	}

	public static void removeTypeFormat(Class<?> clazz) {
		DefaultFormat.CLASS_FORMAT_MAP.remove(clazz);
	}


	public static void createExcelImageSaveFunction(ImageFileSaveFunction imageFilePath) {
		imageFileFunction = imageFilePath;
	}


	public static ImageFileSaveFunction getImageFileFunction() {
		return Optional.ofNullable(imageFileFunction).orElse((imageFilePath, type) -> "");
	}
	
	/**
	 * 设置默认的 Excel 读取引擎
	 * @param engine 引擎版本（V2 或 V3）
	 */
	public static void setDefaultReadEngine(ReadEngine engine) {
		defaultReadEngine = engine;
	}
	
	/**
	 * 获取默认的 Excel 读取引擎
	 * @return 当前默认引擎
	 */
	public static ReadEngine getDefaultReadEngine() {
		return defaultReadEngine;
	}
}
