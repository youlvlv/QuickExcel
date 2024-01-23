package com.xyf.excel.model;

import com.xyf.excel.core.ColumnExcelCore;
import com.xyf.excel.core.ExcelUtil;
import com.xyf.excel.core.RowExcelCore;
import com.xyf.excel.entity.ExcelEntity;

import java.util.List;

public class ExcelBaseModel {
	protected OperationalModel operationalModel = OperationalModel.ROW;

	protected static RowExcelCore rowCore = new RowExcelCore();
	protected static ColumnExcelCore columnCore = new ColumnExcelCore();

	protected ExcelUtil util(){
		switch (operationalModel){
			case COLUMN:
				return columnCore;
			case ROW:
			default:
				return rowCore;
		}
	}

	public void switchOperational(OperationalModel operationalModel) {
		this.operationalModel = operationalModel;
	}

	public OperationalModel getOperationalModel() {
		return operationalModel;
	}

	public static <T> List<ExcelEntity> getExcelEntities(Class<T> entity){
		return rowCore.getExcelEntities(entity);
	}
}
