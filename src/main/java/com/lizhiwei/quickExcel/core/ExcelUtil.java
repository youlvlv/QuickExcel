package com.lizhiwei.quickExcel.core;


import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.*;
import com.lizhiwei.quickExcel.exception.ExcelValueError;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormat;
import com.lizhiwei.quickExcel.model.MoreRowModel;
import com.lizhiwei.quickExcel.model.RowModel;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 核心算法类
 *
 * @author lizhiwei
 */
public abstract class ExcelUtil {

	/**
	 * 序号
	 */
	private static final ExcelEntity index = new ExcelEntity(ParamType.INDEX);

	/**
	 * 直接获取属性值
	 * @param excelEntity
	 * @param t
	 * @return
	 * @param <T>
	 * @throws NoSuchFieldException
	 * @throws IllegalAccessException
	 */
	protected static <T> String getParamString(ExcelEntity excelEntity, T t) throws NoSuchFieldException, IllegalAccessException {
		Field field;
		String str = excelEntity.getProperty();
		//获取该属性
		field = t.getClass().getDeclaredField(str);
		field.setAccessible(true);
		Object o = field.get(t);
		String value = "";
		ExcelFormat format = excelEntity.getFormat();
		value = format.WriterToExcel(o);
		return value;
	}

	/**
	 * 根据 实体类生成 excel实体类
	 *
	 * @param entity
	 * @param <T>
	 * @return
	 */
	public <T> List<ExcelEntity> getExcelEntities(Class<T> entity) {
		return this.getExcelEntities(entity, false, null);
	}


	/**
	 * 根据 实体类生成 excel实体类
	 *
	 * @param entity
	 * @param <T>
	 * @return
	 */
	public <T> List<ExcelEntity> getExcelEntities(Class<T> entity, boolean hasIndex, IndexType type) {
		Field[] fields = entity.getDeclaredFields();
		List<ExcelEntity> listTitle = new ArrayList<>();

		// 转换器缓存,默认初始化默认构造器
		Map<Class<?>, ExcelFormat<?>> formatCache = ExcelConfig.getFormatCache();

		//检查所有的属性
		for (Field field : fields) {
			//设置属性默认可访问，防止private阻止访问
			field.setAccessible(true);
			//判断是否包含Excel注解
			if (field.isAnnotationPresent(Excel.class)) {
				//获取Excel注解
				Excel e = field.getDeclaredAnnotation(Excel.class);
				extractedExcelFormat(formatCache, e.format());
				ExcelEntity excelEntity;
				if (e.type() == ParamType.METHOD) {
					String get = "get" + Pattern.compile("^.").matcher(field.getName()).replaceFirst(m -> m.group().toUpperCase());
					String set = "set" + Pattern.compile("^.").matcher(field.getName()).replaceFirst(m -> m.group().toUpperCase());
					try {
						entity.getMethod(get);
						entity.getMethod(set, field.getType());
					} catch (NoSuchMethodException ex) {
						throw new RuntimeException("未找到get/set方法");
					}
				}
				//构造excel实体类
				excelEntity = new ExcelEntity(e, formatCache.get(e.format()), field.getName(), field.getType());
				listTitle.add(excelEntity);
			}
		}
		//判断当前是否有自主排序
		if (listTitle.stream().anyMatch(x -> x.getIndex() != -1)) {
			listTitle.sort(Comparator.comparingInt(ExcelEntity::getIndex));
		}
		//判断当前是否有序号行
		if (hasIndex) {
			if (type == IndexType.FINALLY) {
				listTitle.add(index);
			} else {
				listTitle.add(0, index);
			}
		}
		//重置排序
		AtomicInteger i = new AtomicInteger();
		listTitle.forEach(x -> x.setIndex(i.getAndIncrement()));
		return listTitle;
	}

	/**
	 * 生成excel转换器
	 *
	 * @param formatCache 转换器缓存
	 * @param format      转换器
	 */
	private static void extractedExcelFormat(Map<Class<?>, ExcelFormat<?>> formatCache, Class<? extends ExcelFormat> format) {
		//判断当前转换器是否存在缓存
		if (!formatCache.containsKey(format)) {
			//不存在缓存，则进行实例化
			ExcelFormat<?> excelFormat;
			try {
				excelFormat = format.getDeclaredConstructor().newInstance();
			} catch (InvocationTargetException | InstantiationException | IllegalAccessException |
			         NoSuchMethodException ex) {
				//实例化失败，则使用默认的转换器
				ex.printStackTrace();
				excelFormat = new DefaultFormat();
			}
			//放入缓存列表中
			formatCache.put(format, excelFormat);
		} else {
			//执行重新初始化命令
			formatCache.put(format,formatCache.get(format).init());
		}
	}

	/**
	 * 判断当前是否为get方法
	 *
	 * @param method
	 * @return
	 */
	public static boolean isGetter(Method method) {
		if (!method.getName().startsWith("get")) {
			return false;
		}
		//get方法肯定没有参数
		if (method.getParameterTypes().length != 0) {
			return false;
		}
		return !void.class.equals(method.getReturnType());
	}

	/**
	 * 生成数据头
	 *
	 * @param sheet
	 * @param listTitle
	 * @return
	 */
	public SheetModel setSheetHeader(SheetModel sheet, List<ExcelEntity> listTitle) {
		//创建表格的样式
		CellStyle cs = sheet.getExcel().getWorkbook().createCellStyle();
		//设置水平、垂直居中
		cs.setAlignment(HorizontalAlignment.CENTER);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		//设置字体
		Font headerFont = sheet.getExcel().getWorkbook().createFont();
		headerFont.setFontHeightInPoints((short) 12);
		/*headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);*/
		headerFont.setBold(true);
		headerFont.setFontName("宋体");
		cs.setFont(headerFont);
		cs.setWrapText(true);//是否自动换行
		return this.setSheetHeader(sheet, listTitle, cs, (short) 0);
	}

	/**
	 * 生成数据头
	 *
	 * @param sheet
	 * @param listTitle
	 * @return
	 */
	public SheetModel setSheetHeader(SheetModel sheet, List<ExcelEntity> listTitle, CellStyle cs) {
		return this.setSheetHeader(sheet, listTitle, cs, (short) 0);
	}

	/**
	 * 生成数据头
	 *
	 * @param sheet
	 * @param listTitle
	 * @return
	 */
	public SheetModel setSheetHeader(SheetModel sheet, List<ExcelEntity> listTitle, CellStyle cs, short s) {
		//去掉所有禁止导出的字段
		AtomicInteger i = new AtomicInteger();
		listTitle = listTitle.stream().filter(ExcelEntity::isWrite).collect(Collectors.toList());
		listTitle.forEach(x -> x.setIndex(i.getAndIncrement()));
		//判断是否有多行头
		boolean moreRow = listTitle.stream().filter(x -> !x.getTopName().isEmpty()).findAny().orElse(null) != null;
		if (moreRow) {
			MoreRowModel xRow0 = sheet.newMoreRow();
			//获取所有非默认头的字段
			Map<String, List<ExcelEntity>> group = listTitle.stream().filter(x -> !x.getTopName().isEmpty())
					.collect(Collectors.groupingBy(ExcelEntity::getTopName));
//            Map<Integer,TopName> type = new HashMap<>();
			group.forEach((k, v) -> {
				if (v.size() > 1) {
					xRow0.setHeaderValue(v.get(0).getIndex(), v.get(0).getIndex() + v.size() - 1, k, cs);
					xRow0.setSecondHeaderValue(v, cs);
				} else {
					xRow0.setValue(v.get(0).getIndex(), k, v.get(0).getTitle(), cs);
//					xRow0.setSecondHeaderValue(v, cs);
				}
			});
			for (ExcelEntity excelEntity : listTitle) {
				if (excelEntity.getTopName().isEmpty()) {
					xRow0.setValue(excelEntity.getIndex(), excelEntity.getTitle(), cs);
				}
			}
		} else {
			//创建一行
			RowModel xRow0 = sheet.newRow();
			// 判断是否存在行高
			if (s == 0) {
				for (ExcelEntity excelEntity : listTitle) {
					xRow0.setValue(excelEntity.getIndex(), excelEntity.getTitle(), cs);
				}
			} else {
				for (ExcelEntity excelEntity : listTitle) {
					xRow0.setValue(excelEntity.getIndex(), excelEntity.getTitle(), cs, s);
				}
			}
		}
		for (ExcelEntity excelEntity : listTitle) {
			sheet.getSheet().setColumnWidth(excelEntity.getIndex(), excelEntity.getWidth());
		}
		return sheet;
	}



	public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle) {
		return this.setSheetContent(sheet, listContent, listTitle, null);
	}


	/**
	 * 配置(赋值)表格内容部分
	 *
	 * @param sheet       工作表
	 * @param listContent 内容
	 * @param listTitle   表头
	 * @param since       合并表格
	 */
	public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since) {
		//创建内容样式（头部以下的样式）
		CellStyle cs = sheet.getExcel().getWorkbook().createCellStyle();
		cs.setWrapText(true);
		//设置水平垂直居中
		cs.setAlignment(HorizontalAlignment.CENTER);
		cs.setVerticalAlignment(VerticalAlignment.CENTER);
		return this.setSheetContent(sheet, listContent, listTitle, since, cs, (short) 0);
	}

	/**
	 * 配置(赋值)表格内容部分
	 *
	 * @param sheet       工作表
	 * @param listContent 内容
	 * @param listTitle   表头
	 * @param since       合并表格
	 * @param cs          样式
	 */
	public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since, CellStyle cs) {
		return this.setSheetContent(sheet, listContent, listTitle, since, cs, (short) 0);
	}


	/**
	 * 配置(赋值)表格内容部分
	 *
	 * @param sheet       工作表
	 * @param listContent 内容
	 * @param listTitle   表头
	 * @param since       合并表格
	 * @param cs          样式
	 * @param ss          行高
	 * @param <T>         实体类
	 * @return 返回
	 */
	public abstract  <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since, CellStyle cs, short ss);
}
