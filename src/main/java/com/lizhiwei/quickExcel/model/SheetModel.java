package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.entity.Since;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
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
	public <T> SheetModel createContent(Class<T> entity, List<T> listContent, CellStyle style, short s) {
		List<ExcelEntity> list = getEntities(entity);
		return util().setSheetContent(this, listContent, list, null, style, s);
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


	/**
	 * 创建画图 Helper
	 */
	public ClientAnchor createHelper() {
		return xSheet.getWorkbook().getCreationHelper().createClientAnchor();
	}

	/**
	 * 添加图片
	 * @param imagePath 图片路径
	 * @param sheet  工作表
	 * @return 图片索引
	 */
	public static void addPicture(String imagePath, Sheet sheet) throws IOException {
		InputStream inputStream = new FileInputStream(imagePath);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		inputStream.close();

		Workbook workbook = sheet.getWorkbook(); // 获取工作簿实例
		int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG); // 根据图片类型添加到工作簿

		CreationHelper createHelper = workbook.getCreationHelper();
		ClientAnchor anchor = createHelper.createClientAnchor();

		// 设置图片在单元格的位置，例如图片从A1单元格开始
		anchor.setCol1(0);
		anchor.setRow1(0);
		// 可以根据需要设置其他位置参数

		Drawing<?> drawing = sheet.createDrawingPatriarch(); // 创建绘图 patriarch 对象
		Picture pict = drawing.createPicture(anchor, pictureIdx); // 在指定位置创建图片

		// 如果需要统一设置图片宽度，可以进行如下操作：
		float scale = 1; // 定义缩放比例
		int widthPx = 42; // 图片的像素宽度
		pict.resize(widthPx * Units.EMU_PER_PIXEL / scale, -1); // -1 表示自动计算高度以保持纵横比
	}

}
