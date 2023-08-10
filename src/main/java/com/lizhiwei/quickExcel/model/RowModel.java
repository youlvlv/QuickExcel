package com.lizhiwei.quickExcel.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class RowModel {
	/**
	 * 当前行数
	 */
	protected final int rowNumber;
	protected final XSSFRow row;
	protected final SheetModel sheet;
	/**
	 * 当前单元格位置
	 */
	protected Integer order = 0;


	public RowModel(int rowNumber, XSSFRow row, SheetModel sheetModel) {
		this.rowNumber = rowNumber;
		this.row = row;
		this.sheet = sheetModel;
	}

	protected XSSFCell createCell(int i) {
		return row.createCell(i);
	}

	/**
	 * 设置合并单元格
	 *
	 * @return
	 */
	public RowModel setMergerValue(int start, int end, String value, CellStyle style) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, start, end));
		XSSFCell cell = createCell(start);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		return this;
	}

	/**
	 * 设置合并单元格(添加样式)
	 *
	 * @return
	 */
	public RowModel setMergerValue(int start, int end, String value, CellStyle style, short s) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, start, end));
		XSSFCell cell = createCell(start);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		cell.getRow().setHeight(s);
		return this;
	}

	/**
	 * 设置合并单元格
	 *
	 * @return
	 */
	public RowModel setMergerValue(int start, int end, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, start, end));
		XSSFCell cell = createCell(start);
		cell.setCellValue(value);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		return this;
	}

	public SheetModel over() {
		return sheet;
	}

	/**
	 * 设置值
	 *
	 * @param i
	 * @param value
	 * @return
	 */
	public RowModel setValue(int i, String value, CellStyle style) {
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
	public RowModel setValue(int i, String value, CellStyle style, short s) {
		XSSFCell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		cell.getRow().setHeight(s);
		return this;
	}

	/**
	 * 设置值
	 *
	 * @param i     第几列表
	 * @param value 内容
	 * @return 返回
	 */
	public RowModel setValue(int i, String value) {
		XSSFCell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		return this;
	}

	/**
	 * 设置值
	 *
	 * @param value 内容
	 * @return 返回
	 */
	public RowModel setValue(String value) {
		XSSFCell cell = createCell(order++);
		cell.setCellValue(value);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		return this;
	}
}
