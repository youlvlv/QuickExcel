package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.util.List;
import java.util.function.Function;

public class MoreRowModel extends RowBaseModel<MoreRowModel> {

	/**
	 * 结束行数
	 */
	protected final int endRowNumber;

	protected final Row secondRow;

	protected final Sheet xSheet;


	public MoreRowModel(int rowNumber, int endRowNumber, Row row, Sheet xSheet, SheetModel sheetModel) {
		super(rowNumber, row, sheetModel);
		super.chain = this;
		this.endRowNumber = Integer.parseInt(String.valueOf(endRowNumber));
		this.xSheet = xSheet;
		this.secondRow = xSheet.createRow(rowNumber+1);
	}

	public RowModelSeparation getRow(int i) {
		if (i == 0) {
			return new RowModelSeparation(rowNumber, row, sheet, this);
		} else if (i > 0 && i + rowNumber <= endRowNumber) {
			return new RowModelSeparation(i + 1,xSheet.createRow(i + rowNumber) , sheet, this);
		} else {
			throw new RuntimeException("超出范围");
		}

	}

	@Override
	public MoreRowModel setValue(int i, String value, CellStyle style, short s) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		Cell cell2 = secondRow.createCell(i);
		cell2.setCellValue("");
		cell2.setCellStyle(style);
		return super.setValue(i, value, style, s);
	}

	@Override
	public MoreRowModel setValue(int i, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		Cell cell2 = secondRow.createCell(i);
		cell2.setCellValue("");
		cell2.setCellStyle(sheet.getExcel().getDefaultStyle());
		return super.setValue(i, value);
	}

	@Override
	public MoreRowModel setValue(int i, String filePath, SheetModel sheetModel) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		return super.setValue(i, filePath, sheetModel);
	}

	@Override
	public MoreRowModel setValue(int i, List<String> filePath, SheetModel sheetModel) throws IOException {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		return super.setValue(i, filePath, sheetModel);
	}

	@Override
	public MoreRowModel setValue(String value) {

		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, order, order));
		Cell cell2 = secondRow.createCell(order);
		cell2.setCellValue("");
		cell2.setCellStyle(sheet.getExcel().getDefaultStyle());
		return super.setValue(value);
	}

	@Override
	public MoreRowModel setValue(String value, Function<CellStyle, CellStyle> style) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, order, order));
		return super.setValue(value, style);
	}

	public MoreRowModel setValue(int i, String firstValue, String secondValue, CellStyle style) {
		Cell cell = row.createCell(i);
		cell.setCellValue(firstValue);
		cell.setCellStyle(style);
		Cell cell2 = secondRow.createCell(i);
		cell2.setCellValue(secondValue);
		cell2.setCellStyle(style);
		return this;
	}

	public MoreRowModel setHeaderValue(int i, int end, String value, CellStyle cs) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, i, end));
		for (int j = i; j <= end; j++) {
			super.setValue(j, "", cs);
		}
		super.setValue(i, value, cs);
		return this;
	}

	public MoreRowModel setValue(int i,int rowNumber, String value, CellStyle style) {
		if (value.contains("DRAW_IMAGE::")) {
			setValue(i, value.replace("DRAW_IMAGE::", ""), this.sheet);
		}
		Cell cell = createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		Cell cell2 = secondRow.createCell(i);
		//cell2.setCellValue("");
		cell2.setCellStyle(style);
		return chain;
	}

	public MoreRowModel setSecondValue(int i, String value, CellStyle style) {
		if (value.contains("DRAW_IMAGE::")) {
			setValue(i, value.replace("DRAW_IMAGE::", ""), this.sheet);
		}
		Cell cell = secondRow.createCell(i);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		return chain;
	}

	@Override
	public MoreRowModel setValue(int i, String value, CellStyle style) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		super.setValue(i, value, style);
		Cell cell2 = secondRow.createCell(i);
		//cell2.setCellValue("");
		cell2.setCellStyle(style);
		return this;
	}

	@Override
	public MoreRowModel setMergerValue(int start, int end, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, start, end));
		super.setValue(start, value, sheet.getExcel().getDefaultStyle());
		Cell cell2 = secondRow.createCell(start);
		cell2.setCellValue("");
		cell2.setCellStyle( sheet.getExcel().getDefaultStyle());
		for (int i = ++start; i <= end; i++) {
			super.setValue(i, "", sheet.getExcel().getDefaultStyle());
			this.setSecondValue(i, "", sheet.getExcel().getDefaultStyle());
		}
		return this;
	}


	public void setSecondHeaderValue(List<ExcelEntity> v, CellStyle style) {
		for (ExcelEntity excelEntity : v) {
			Cell cell = secondRow.createCell(excelEntity.getIndex());
			cell.setCellValue(excelEntity.getTitle());
			cell.setCellStyle(style);
		}
	}

	public static class RowModelSeparation extends RowBaseModel<RowModelSeparation> {

		protected final MoreRowModel moreRowModel;

		public RowModelSeparation(int rowNumber, Row row, SheetModel sheetModel, MoreRowModel moreRowModel) {
			super(rowNumber, row, sheetModel);
			super.chain = this;
			this.moreRowModel = moreRowModel;
			// 拷贝 moreRowModel.order
			this.order = Integer.valueOf(moreRowModel.order.toString());
		}

		public MoreRowModel overSignRow() {
			return moreRowModel;
		}

		public MoreRowModel overSignRowAndAddOrder() {
			moreRowModel.order = this.order;
			return moreRowModel;
		}
	}
}
