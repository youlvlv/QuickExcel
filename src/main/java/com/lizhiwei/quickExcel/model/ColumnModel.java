package com.lizhiwei.quickExcel.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.List;

/**
 * 列模型
 */
public class ColumnModel {
	// 列号
	protected final int columnNum;

	private final int startRowNum;
	// 行号
	protected Integer rowNum;
	// sheetmodel
	protected final SheetModel sheetModel;


	public ColumnModel(int columnNum, int rowNum, SheetModel sheetModel) {
		this.columnNum = columnNum;
		this.startRowNum = rowNum;
		this.rowNum = Integer.valueOf(rowNum);
		this.sheetModel = sheetModel;
	}


	private XSSFCell createCell() {
		return createCell(rowNum++);
	}

	private XSSFCell createCell(int i) {
		return sheetModel.getSheet().createRow(i).createCell(columnNum);
	}

	public SheetModel over() {
		return sheetModel;
	}

	/**
	 * 设置值
	 *
	 * @param i
	 * @param value
	 * @return
	 */
	public ColumnModel setValue(int i, String value, CellStyle style) {
		XSSFCell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		return this;
	}


	/**
	 * 设置值
	 *
	 * @param i     第几列表
	 * @param value 内容
	 * @param style 样式
	 * @param s     行高
	 * @return 返回
	 */
	public ColumnModel setValue(int i, String value, CellStyle style, short s) {
		XSSFCell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		cell.getRow().setHeight(s);
		return this;
	}

	public ColumnModel setValue(int i, String value) {
		XSSFCell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(sheetModel.getExcel().getDefaultStyle());
		return this;
	}

	/**
	 * 合并单元格
	 *
	 * @param values 所有值
	 * @param style  单元格样式
	 * @param merger 是否存在合并项
	 * @return
	 */
	public ColumnModel setValues(List<String> values, CellStyle style, boolean merger) {
		values.forEach(x -> {
			setValue(x, style);
		});
		if (merger) {
			for (int i = 0; i < values.size(); i++) {
				String one = values.get(i);
				int size = 0;
				for (String value : values) {
					if (one.equals(value)) {
						size++;
					} else {
						break;
					}
				}
				if (size > 0) {
					i += size;
					sheetModel.addMergedRegion(new CellRangeAddress(i, size, columnNum, columnNum));
				}
			}
		}
		return this;
	}


	/**
	 * 设置值
	 *
	 * @param value
	 * @return
	 */
	public ColumnModel setValue(String value, CellStyle style) {
		XSSFCell cell = createCell();
		cell.setCellValue(value);
		cell.setCellStyle(style);
		return this;
	}


	/**
	 * 设置值
	 *
	 * @param value 内容
	 * @param style 样式
	 * @param s     行高
	 * @return 返回
	 */
	public ColumnModel setValue(String value, CellStyle style, short s) {
		XSSFCell cell = createCell();
		cell.setCellValue(value);
		cell.setCellStyle(style);
		cell.getRow().setHeight(s);
		return this;
	}

	public ColumnModel setValue(String value) {
		XSSFCell cell = createCell();
		cell.setCellValue(value);
		cell.setCellStyle(sheetModel.getExcel().getDefaultStyle());
		return this;
	}
}
