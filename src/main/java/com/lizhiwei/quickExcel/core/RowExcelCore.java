package com.lizhiwei.quickExcel.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.Since;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;
import com.lizhiwei.quickExcel.model.RowModel;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 行模式运行下的算法
 */
public class RowExcelCore extends ExcelUtil {



	@Override
    public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle,
                                          List<Since> since, CellStyle cs, short ss, BiConsumer<T, RowModel> row) {
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
					for (ExcelEntity excelEntity : listTitle) {
						switch (excelEntity.getParamType()) {
							case INDEX: {
                                xRow.setValue(String.valueOf(sheet.getNum()), cs);
								break;
							}
							// 属性
							case FIELD: {
								String value = getParamString(excelEntity, t);
								//循环设置每列的值
                                xRow.setValue(value, cs);
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
								ExcelFormatBase format = excelEntity.getFormat();
								value = format.WriterToExcel(o);
								//循环设置每列的值
                                xRow.setValue(value, cs);
								break;
							}
						}
					}
					if (row != null) {
						row.accept(t, xRow);
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
				throw new ExcelValueException(e);
			}
		}
		return sheet;
	}
}
