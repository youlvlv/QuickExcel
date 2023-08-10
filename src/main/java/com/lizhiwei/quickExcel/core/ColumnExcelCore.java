package com.lizhiwei.quickExcel.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.Since;
import com.lizhiwei.quickExcel.exception.ExcelValueError;
import com.lizhiwei.quickExcel.format.ExcelFormat;
import com.lizhiwei.quickExcel.model.ColumnModel;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 列模式运行下的算法
 */
public class ColumnExcelCore extends ExcelUtil {


	@Override
	public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since, CellStyle cs, short ss) {
		//去掉所有禁止导出的字段
		listTitle = listTitle.stream().filter(ExcelEntity::isWrite).collect(Collectors.toList());
		int start = sheet.getRowNum();
		if (null != listContent && listContent.size() > 0) {
			try {
				for (ExcelEntity excelEntity : listTitle) {
					boolean merger = since.stream().anyMatch(x -> excelEntity.getProperty().equals(x.getTitle()));
					ColumnModel column = sheet.newColumn();
					//获取类属性
					Field field;
					Method getter;
					List<String> values = new ArrayList<>();
					for (T t : listContent) {
						switch (excelEntity.getParamType()) {
							case INDEX: {
								column.setValue(String.valueOf(sheet.getNum()), cs);
								break;
							}
							// 属性
							case FIELD: {
								String value = getParamString(excelEntity, t);
								//循环设置每列的值
								values.add(value);
								break;
							}
							// 方法
							case METHOD: {
								String str = excelEntity.getProperty();
								String get = "get" + Pattern.compile("^.").matcher(str).replaceFirst(m -> m.group().toUpperCase());
								//获取该属性
								getter = t.getClass().getMethod(get);
								Object o = getter.invoke(t);
								String value = "";
								ExcelFormat format = excelEntity.getFormat();
								values.add(format.WriterToExcel(o));
								//循环设置每列的值
								break;
							}
						}
					}
					column.setValues(values, cs, merger);
				}
			} catch (IllegalAccessException | NoSuchFieldException | NoSuchMethodException |
			         InvocationTargetException e) {
				throw new ExcelValueError(e);
			}
		}
		return sheet;
	}


}
