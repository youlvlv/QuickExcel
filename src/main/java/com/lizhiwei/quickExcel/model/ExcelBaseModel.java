package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.core.ColumnExcelCore;
import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.core.RowExcelCore;
import com.lizhiwei.quickExcel.entity.ExcelEntity;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelBaseModel {
	protected OperationalModel operationalModel = OperationalModel.ROW;

	protected static RowExcelCore rowCore = new RowExcelCore();
	protected static ColumnExcelCore columnCore = new ColumnExcelCore();

	/**
	 * MIME类型和后缀对应关系
	 */
	protected static final Map<String, String> MIME_TYPE_TO_EXTENSION = new HashMap<>() {{
		put("image/jpeg", ".jpg");
		put("image/png", ".png");
		put("image/gif", ".gif");
		put("image/bmp", ".bmp");
		put("image/x-emf", ".emf");
		// 添加其他需要支持的MIME类型和后缀对应关系
	}};

	public static <T> List<ExcelEntity> getExcelEntities(Class<T> entity) {
		return rowCore.getExcelEntities(entity);
	}

	public void switchOperational(OperationalModel operationalModel) {
		this.operationalModel = operationalModel;
	}

	public OperationalModel getOperationalModel() {
		return operationalModel;
	}

	protected ExcelUtil util() {
        return switch (operationalModel) {
            case COLUMN -> columnCore;
            default -> rowCore;
        };
	}
}
