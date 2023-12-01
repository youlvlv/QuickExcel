package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.entity.Since;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetModel extends ExcelBaseModel {
	private final XSSFSheet xSheet;
	private final ExcelModel excel;

	private IndexType type;
	/**
	 * 序号
	 */
	private int num = 1;
	/**
	 * 行号
	 */
	protected int rowNum = 0;
	/**
	 * 列号
	 */
	protected int columnNum = 0;

	protected int overRowNum;
	/**
	 * 运行模式
	 */
	protected OperationalModel operationalModel = OperationalModel.ROW;

	public SheetModel(XSSFSheet xSheet, ExcelModel excel) {
		this.xSheet = xSheet;
		this.excel = excel;
	}

	/**
	 * 创建 数据信息
	 *
	 * @param entity
	 * @param listContent
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createInfo(Class<T> entity, List<T> listContent) {
		List<ExcelEntity> list = util().getExcelEntities(entity);
		SheetModel newSheet = util().setSheetHeader(this, list);
		return util().setSheetContent(newSheet, listContent, list, null);
	}

	public SheetModel createSerialNumber(IndexType type) {
		this.type = type;
		return this;
	}

	/**
	 * 创建数据头
	 *
	 * @param entity
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createHeader(Class<T> entity) {
		List<ExcelEntity> list = getEntities(entity);
		return util().setSheetHeader(this, list);
	}

	/**
	 * 创建表头带样式的
	 *
	 * @param entity
	 * @param cellStyle
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createHeader(Class<T> entity, CellStyle cellStyle, short s) {
		List<ExcelEntity> list = getEntities(entity);
		return util().setSheetHeader(this, list, cellStyle, s);
	}

	/**
	 * 录入数据信息
	 *
	 * @param entity  实体类class
	 * @param content 数据
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(Class<T> entity, T content) {
		List<ExcelEntity> list = getEntities(entity);
		List<T> first = new ArrayList<>();
		first.add(content);
		return util().setSheetContent(this, first, list, null);
	}


	/**
	 * 录入数据信息
	 *
	 * @param entity      实体类class
	 * @param listContent 数据
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(Class<T> entity, List<T> listContent) {
		List<ExcelEntity> list = getEntities(entity);
		return util().setSheetContent(this, listContent, list, null);
	}

    /**
     * 录入数据信息
     *
     * @param entity      实体类class
     * @param listContent 数据
     * @param <T>
     * @return
     */
    public <T> SheetModel createContent(Class<T> entity, List<T> listContent,CellStyle style,short s) {
        List<ExcelEntity> list = getEntities(entity);
        return util().setSheetContent(this, listContent, list, null,style,s);
    }


	private <T> List<ExcelEntity> getEntities(Class<T> entity) {
		List<ExcelEntity> list;
		if (type != null) {
			list = util().getExcelEntities(entity, true, type);
		} else {
			list = util().getExcelEntities(entity);
		}
		return list;
	}

	/**
	 * 录入数据信息
	 *
	 * @param entity      实体类class
	 * @param listContent 数据
	 * @param since       合并
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(Class<T> entity, List<T> listContent, Since... since) {
		SheetModel sheetModel = null;
		if (operationalModel == OperationalModel.ROW) {
			List<ExcelEntity> list = getEntities(entity);
			sheetModel = util().setSheetContent(this, listContent, list, Arrays.asList(since));
		} else {

		}
		return sheetModel;
	}

	/**
	 * 获取单行数据
	 *
	 * @return
	 */
	public RowModel newRow() {
		return new RowModel(rowNum, xSheet.createRow(rowNum++), this);
	}

	/**
	 * 获取单列信息
	 *
	 * @return
	 */
	public ColumnModel newColumn() {
		return new ColumnModel(columnNum, rowNum, this);
	}

	/**
	 * 新生成多行
	 *
	 * @return
	 */
	public MoreRowModel newMoreRow() {
		return new MoreRowModel(rowNum, rowNum + 1, xSheet.createRow(rowNum++), xSheet.createRow(rowNum++), this);
	}

	public XSSFSheet getSheet() {
		return xSheet;
	}

	public ExcelModel getExcel() {
		return excel;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void addRowNum() {
		rowNum++;
	}


	/**
	 * 结束本sheet编辑
	 *
	 * @return
	 */
	public ExcelModel over() {
		return excel;
	}

	public void addMergedRegion(CellRangeAddress cellRangeAddress) {
		xSheet.addMergedRegion(cellRangeAddress);
	}

	public int getNum() {
		return num++;
	}

	/**
	 * 创建一共绘图对象
	 */
	public Drawing createDrawingPatriarch() {
		return xSheet.createDrawingPatriarch();
	}
}
