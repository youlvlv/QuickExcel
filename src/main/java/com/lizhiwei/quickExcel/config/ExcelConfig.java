package com.lizhiwei.quickExcel.config;

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
	private static BiFunction<InputStream, String, String> imageFileFunction;

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


	public static void createExcelImageSaveFunction(BiFunction<InputStream, String, String> imageFilePath) {
		imageFileFunction = imageFilePath;
	}


	public static BiFunction<InputStream, String, String> getImageFileFunction() {
		return Optional.ofNullable(imageFileFunction).orElse((imageFilePath, type) -> "");
	}
}
