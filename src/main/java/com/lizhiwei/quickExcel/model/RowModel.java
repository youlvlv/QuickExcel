package com.lizhiwei.quickExcel.model;

import org.apache.poi.ss.usermodel.Row;

public class RowModel extends RowBaseModel<RowModel> {

	/**
	 * 当前单元格位置
	 */
	protected Integer order = 0;


	public RowModel(int rowNumber, Row row, SheetModel sheetModel) {
		super(rowNumber, row, sheetModel);
		super.chain = this;
	}

}
