package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.util.List;

public class MoreRowModel extends RowBaseModel<MoreRowModel> {

	/**
	 * 结束行数
	 */
	protected final int endRowNumber;


	protected final Row secondRow;


	public MoreRowModel(int rowNumber, int endRowNumber, Row row, Row secondRow, SheetModel sheetModel) {
		super(rowNumber, row, sheetModel);
		super.chain = this;
		this.endRowNumber = endRowNumber;
		this.secondRow = secondRow;
	}

	public RowModelSeparation getRow(int i) {
		if (i == 0) {
			return new RowModelSeparation(rowNumber, row, sheet, this);
		} else if (i > 0 && i + rowNumber <= endRowNumber) {
			return new RowModelSeparation(i + 1, secondRow, sheet, this);
		} else {
			throw new RuntimeException("超出范围");
		}

	}

	@Override
	public MoreRowModel setValue(int i, String value, CellStyle style, short s) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		return super.setValue(i, value, style, s);
	}

	@Override
	public MoreRowModel setValue(int i, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
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
		return super.setValue(value);
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
		super.setValue(i, value, cs);
		return this;
	}

	@Override
	public MoreRowModel setValue(int i, String value, CellStyle style) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
		super.setValue(i, value, style);
		return this;
	}

	@Override
	public MoreRowModel setMergerValue(int start, int end, String value) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, start, end));
		super.setValue(start, value);
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
		}

		public MoreRowModel overSignRow() {
			return moreRowModel;
		}
	}
}
