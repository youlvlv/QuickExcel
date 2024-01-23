package com.xyf.excel.core;

import com.xyf.excel.entity.ExcelEntity;
import com.xyf.excel.entity.Since;
import com.xyf.excel.exception.ExcelValueError;
import com.xyf.excel.format.ExcelFormat;
import com.xyf.excel.model.RowModel;
import com.xyf.excel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 行模式运行下的算法
 */
public class RowExcelCore extends ExcelUtil {



	@Override
	public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since, CellStyle cs, short ss) {
		//去掉所有禁止导出的字段
		listTitle = listTitle.stream().filter(ExcelEntity::isWrite).collect(Collectors.toList());
		for (int i = 0; i < listTitle.size(); i++) {
			listTitle.get(i).setIndex(i);
		}
		int start = sheet.getRowNum();
		if (null != listContent && !listContent.isEmpty()) {
			try {
				for (T t : listContent) {
					RowModel xRow = sheet.newRow();
					//获取类属性
					Field field;
					Method getter;
					int order = 0;
					for (ExcelEntity excelEntity : listTitle) {
						switch (excelEntity.getParamType()) {
							case INDEX: {
								xRow.setValue(order++, String.valueOf(sheet.getNum()), cs);
								break;
							}
							// 属性
							case FIELD: {
								String value = getParamString(excelEntity, t);
								//循环设置每列的值
								xRow.setValue(order++, value, cs);
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
								value = format.WriterToExcel(o);
								//循环设置每列的值
								xRow.setValue(order++, value, cs);
								break;
							}
						}
					}
				}
				if (since != null) {
					for (Since s : since) {
						int i = listTitle.stream().filter(x -> x.getProperty().equals(s.getTitle())).findFirst().get().getIndex();
						sheet.addMergedRegion(new CellRangeAddress(start, sheet.getRowNum() - 1, i, i));
					}
				}

			} catch (IllegalAccessException | NoSuchFieldException | NoSuchMethodException |
			         InvocationTargetException e) {
				throw new ExcelValueError(e);
			}
		}
		return sheet;
	}
}
