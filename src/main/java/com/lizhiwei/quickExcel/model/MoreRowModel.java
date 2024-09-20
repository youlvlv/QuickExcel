package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class MoreRowModel extends RowModel {

	/**
	 * 结束行数
	 */
	protected final int endRowNumber;


	protected final Row secondRow;


	public MoreRowModel(int rowNumber, int endRowNumber, Row row, Row secondRow, SheetModel sheetModel) {
		super(rowNumber, row, sheetModel);
		this.endRowNumber = endRowNumber;
		this.secondRow = secondRow;
	}


	public RowModel setValue(int i, String firstValue, String secondValue, CellStyle style) {
		Cell cell = row.createCell(i);
		cell.setCellValue(firstValue);
		cell.setCellStyle(style);
		Cell cell2 = secondRow.createCell(i);
		cell2.setCellValue(secondValue);
		cell2.setCellStyle(style);
		return this;
	}

	public RowModel setHeaderValue(int i, int end, String value, CellStyle cs) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, i, end));
		Cell cell = row.createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(cs);
		return this;
	}

	@Override
	public RowModel setValue(int i, String value, CellStyle style) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		Cell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		return this;
	}

	@Override
	public RowModel setMergerValue(int start, int end, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, start, end));
		Cell cell = createCell(start);
		cell.setCellValue(value);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		return this;
	}


	public void setSecondHeaderValue(List<ExcelEntity> v, CellStyle style) {
		for (ExcelEntity excelEntity : v) {
			Cell cell = secondRow.createCell(excelEntity.getIndex());
			cell.setCellValue(excelEntity.getTitle());
			cell.setCellStyle(style);
		}
	}
}
