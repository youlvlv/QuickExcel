package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.entity.Since;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.BiConsumer;

public class SheetModel extends ExcelBaseModel implements ModelBase<ExcelModel> {
	private static final Logger log = LogManager.getLogger(SheetModel.class);
	private final Sheet xSheet;
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

	protected boolean isOpen = false;
	/**
	 * 运行模式
	 */
	protected OperationalModel operationalModel = OperationalModel.ROW;

	public SheetModel(Sheet xSheet, ExcelModel excel) {
		this.xSheet = xSheet;
		this.excel = excel;
	}

	public SheetModel(Sheet xSheet, ExcelModel excel, boolean isOpen) {
		this.xSheet = xSheet;
		this.excel = excel;
		if (isOpen) {
			rowNum = xSheet.getLastRowNum();
			this.isOpen = true;
		}
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
	 * 创建数据头
	 *
	 * @param entity
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createHeader(List<ExcelEntity> entity) {
		return util().setSheetHeader(this, entity);
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
	public SheetModel setColWidth(int colNum, int colWidth) {
		xSheet.setColumnWidth(colNum, colWidth);
		return this;
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
    public <T> SheetModel createContent(Class<T> entity, List<T> listContent, BiConsumer<T, RowModel> row) {
        List<ExcelEntity> list = getEntities(entity);
        return util().setSheetContent(this, listContent, list, null, row);
    }

	/**
	 * 录入数据信息
	 *
	 * @param entity      实体类class
	 * @param listContent 数据
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(List<ExcelEntity> entity, List<T> listContent) {
		return util().setSheetContent(this, listContent, entity, null);
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

	/**
	 * 录入数据信息
	 *
	 * @param entity      实体类class
	 * @param listContent 数据
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(Class<T> entity, List<T> listContent, CellStyle style) {
		List<ExcelEntity> list = getEntities(entity);
		return util().setSheetContent(this, listContent, list, null, style);
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
	 * 录入数据信息(可传修改后的列集合)
	 *
	 * @param list      List<ExcelEntity>
	 * @param listContent 数据
	 * @param since       合并
	 * @param <T>
	 * @return
	 */
	public <T> SheetModel createContent(List<ExcelEntity> list, List<T> listContent, Since... since) {
		SheetModel sheetModel = null;
		if (operationalModel == OperationalModel.ROW) {
			sheetModel = util().setSheetContent(this, listContent, list, Arrays.asList(since));
		}
		return sheetModel;
	}

	// 获取merge对象
	public static CellRangeAddress getMergedRegion(Sheet sheet, int rowNum,
	                                               short cellNum) {
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress merged = sheet.getMergedRegion(i);
			if (merged.isInRange(rowNum, cellNum)) {
				return merged;
			}
		}
		return null;
	}

	private static boolean isNewMergedRegion(
			CellRangeAddressWrapper newMergedRegion,
			Set<CellRangeAddressWrapper> mergedRegions) {
		boolean bool = mergedRegions.contains(newMergedRegion);
		return !bool;
	}

	/**
	 * 添加图片
	 *
	 * @param imagePath 图片路径
	 * @param sheet     工作表
	 * @return 图片索引
	 */
	public static void addPicture(String imagePath, Sheet sheet) throws IOException {
		InputStream inputStream = new FileInputStream(imagePath);
		byte[] imageBytes = inputStream.readAllBytes();
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

	public List<RowModel> readRow(int startRow, int endRow) {
		if (!isOpen) {
			throw new ExcelReadException("当前不是打开模式！禁止使用读取相关 API");
		}
		List<RowModel> list = new ArrayList<>();
		for (int i = startRow; i <= endRow; i++) {
			list.add(new RowModel(i, xSheet.getRow(i), this));
		}
		return list;
	}

	/**
	 * 复制到目标 sheet 复制并合并单元格
	 */
	public void copyToSheet(SheetModel goalSheet, int startSrcRow) {
		this.copyToSheet(goalSheet, startSrcRow, xSheet.getLastRowNum());
	}

	/**
	 * 复制到目标 sheet 并合并单元格
	 */
	public void copyToSheet(SheetModel goalSheet, int startSrcRow, int endSrcRow) {
		Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<>();
		int deltaRows = goalSheet.rowNum - startSrcRow; //如果copy到另一个sheet的起始行数不同
		for (int i = startSrcRow; i <= endSrcRow; i++) {
			Row srcRow = xSheet.getRow(i);
			Row destRow = goalSheet.newRow().row;
			for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
				if (j < 0) {
//					log.error(new MarkerManager.Log4jMarker(j + "").addParents(new MarkerManager.Log4jMarker(srcRow.getCell(j).getStringCellValue())));
					goalSheet.rowNum--;
					break;
				}
				Cell oldCell = srcRow.getCell(j); // old cell
				Cell newCell = destRow.getCell(j); // new cell
				if (oldCell != null) {
					if (newCell == null) {
						newCell = destRow.createCell(j);
					}
					copyCell(oldCell, newCell, goalSheet);
					CellRangeAddress mergedRegion = getMergedRegion(xSheet,
							srcRow.getRowNum(), (short) oldCell.getColumnIndex());
					if (mergedRegion != null) {
						CellRangeAddress newMergedRegion = new CellRangeAddress(
								mergedRegion.getFirstRow() + deltaRows,
								mergedRegion.getLastRow() + deltaRows, mergedRegion
								.getFirstColumn(), mergedRegion
								.getLastColumn());
						CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(
								newMergedRegion);
						if (isNewMergedRegion(wrapper, mergedRegions)) {
							mergedRegions.add(wrapper);
							goalSheet.addMergedRegion(wrapper.range);
						}
					}
				}
			}
		}
	}

	/**
	 * 把原来的Sheet中cell（列）的样式和数据类型复制到新的sheet的cell（列）中
	 *
	 * @param oldCell
	 * @param newCell
	 */
	public void copyCell(Cell oldCell, Cell newCell, SheetModel goalSheet) {

		switch (oldCell.getCellType()) {
			case STRING:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case ERROR:
				newCell.setCellErrorValue(oldCell.getErrorCellValue());
				break;
			case FORMULA:
				newCell.setCellFormula(oldCell.getCellFormula());
				break;
			case BLANK:
			default:
				break;
		}
		if (excel.getExcelType().equals(goalSheet.excel.getExcelType())) {
			newCell.getCellStyle().cloneStyleFrom(oldCell.getCellStyle());
		} else {
			copyCellStyle(oldCell.getCellStyle(), newCell, goalSheet.excel.createStyle(), goalSheet);
		}
	}

	private void copyCellStyle(CellStyle sourceCellStyle, Cell newCell, CellStyle targetCellStyle, SheetModel goalSheet) {

		// 设置对齐方式
		targetCellStyle.setAlignment(sourceCellStyle.getAlignment());
		targetCellStyle.setVerticalAlignment(sourceCellStyle.getVerticalAlignment());

		// 设置边框
		targetCellStyle.setBorderBottom(sourceCellStyle.getBorderBottom());
		targetCellStyle.setBorderLeft(sourceCellStyle.getBorderLeft());
		targetCellStyle.setBorderRight(sourceCellStyle.getBorderRight());
		targetCellStyle.setBorderTop(sourceCellStyle.getBorderTop());

		// 设置填充颜色
		targetCellStyle.setFillForegroundColor(sourceCellStyle.getFillForegroundColor());
		targetCellStyle.setFillPattern(sourceCellStyle.getFillPattern());

		// 设置字体
		// 设置字体
		Font sourceFont = sourceCellStyle.getFontIndex() >= 0 ? excel.xWorkbook.getFontAt(sourceCellStyle.getFontIndex()) : null;
		Font targetFont = createEquivalentFont(sourceFont, goalSheet.excel.xWorkbook);
		targetCellStyle.setFont(targetFont);


		// 设置缩放
		targetCellStyle.setDataFormat(sourceCellStyle.getDataFormat());

		// 设置缩放
		targetCellStyle.setShrinkToFit(sourceCellStyle.getShrinkToFit());

		// 设置旋转
		targetCellStyle.setRotation(sourceCellStyle.getRotation());

		// 设置缩进
		targetCellStyle.setIndention(sourceCellStyle.getIndention());

		// 设置自动换行
		targetCellStyle.setWrapText(sourceCellStyle.getWrapText());

		newCell.setCellStyle(targetCellStyle);
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
        return new MoreRowModel(rowNum, rowNum + 1, xSheet.createRow(rowNum++), xSheet, this);
	}

	/**
	 * 新生成多行
	 *
	 * @return
	 */
	public MoreRowModel newMoreRow(int rowSize) {
        return new MoreRowModel(rowNum, rowNum + rowSize, xSheet.createRow(rowNum++), xSheet, this);
	}

	private Font createEquivalentFont(Font sourceFont, Workbook targetWorkbook) {
		if (sourceFont == null) {
			return targetWorkbook.createFont();
		}

		Font targetFont = targetWorkbook.createFont();
		targetFont.setFontName(sourceFont.getFontName());
		targetFont.setFontHeight(sourceFont.getFontHeight());
		targetFont.setColor(sourceFont.getColor());
		targetFont.setBold(sourceFont.getBold());
		targetFont.setItalic(sourceFont.getItalic());
		targetFont.setStrikeout(sourceFont.getStrikeout());
		targetFont.setTypeOffset(sourceFont.getTypeOffset());
		targetFont.setUnderline(sourceFont.getUnderline());
		targetFont.setCharSet(sourceFont.getCharSet());

		return targetFont;
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

	public Sheet getSheet() {
		return xSheet;
	}

	public void addMergedRegion(CellRangeAddress cellRangeAddress) {
		// 检查新区域是否与现有区域重叠
		boolean isOverlapping = false;
		for (int i = 0; i < xSheet.getNumMergedRegions(); i++) {
			CellRangeAddress existingRegion = xSheet.getMergedRegion(i);
			if (isOverlapping(cellRangeAddress, existingRegion)) {
				isOverlapping = true;
				break;
			}
		}

		if (isOverlapping) {
			// 如果有重叠，先删除旧的合并区域
			for (int i = 0; i < xSheet.getNumMergedRegions(); i++) {
				CellRangeAddress existingRegion = xSheet.getMergedRegion(i);
				if (isOverlapping(cellRangeAddress, existingRegion)) {
					xSheet.removeMergedRegion(i);
					// 删除后需要向前移动索引
					i--;
				}
			}
		}

		// 添加新的合并区域
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

	// 检查两个区域是否重叠
	private boolean isOverlapping(CellRangeAddress newRegion, CellRangeAddress existingRegion) {
		int firstRowNew = newRegion.getFirstRow();
		int lastRowNew = newRegion.getLastRow();
		int firstColNew = newRegion.getFirstColumn();
		int lastColNew = newRegion.getLastColumn();

		int firstRowExisting = existingRegion.getFirstRow();
		int lastRowExisting = existingRegion.getLastRow();
		int firstColExisting = existingRegion.getFirstColumn();
		int lastColExisting = existingRegion.getLastColumn();

		// 检查行和列是否有交集
		return (firstRowNew <= lastRowExisting && lastRowNew >= firstRowExisting) &&
				(firstColNew <= lastColExisting && lastColNew >= firstColExisting);
	}

}


class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {

	public CellRangeAddress range;

	public CellRangeAddressWrapper(CellRangeAddress theRange) {
		this.range = theRange;
	}

	public int compareTo(CellRangeAddressWrapper craw) {
		if (range.getFirstColumn() < craw.range.getFirstColumn()
				&& range.getFirstRow() < craw.range.getFirstRow()) {
			return -1;
		} else if (range.getFirstColumn() == craw.range.getFirstColumn()
				&& range.getFirstRow() == craw.range.getFirstRow()) {
			return 0;
		} else {
			return 1;
		}
	}

}
