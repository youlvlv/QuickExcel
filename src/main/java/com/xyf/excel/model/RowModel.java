package com.xyf.excel.model;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRow;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;

import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE;
import static org.apache.poi.ss.usermodel.ClientAnchor.AnchorType.MOVE_AND_RESIZE;

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
	 * 当前单元格插入图片
	 *
	 * @param i          第几列
	 * @param filePath   路径
	 * @param sheetModel 工作表类
	 */
	public RowModel setValue(int i, String filePath, SheetModel sheetModel) throws IOException {
		XSSFCell cell = createCell(i);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		addPicture(filePath, sheetModel, i, cell.getRowIndex(), cell);
		return this;
	}

	/**
	 * 当前单元格插入多张图片
	 *
	 * @param i          第几列
	 * @param filePath   路径
	 * @param sheetModel 工作表类
	 */
	public RowModel setValue(int i, List<String> filePath, SheetModel sheetModel) throws IOException {
		XSSFCell cell = createCell(i);
		cell.setCellStyle(sheet.getExcel().getDefaultStyle());
		addPicture(filePath, sheetModel, i, cell.getRowIndex(), cell);
		return this;
	}

	/**
	 * @param imagePath  图片路径
	 * @param sheetModel 工作表
	 * @param col1       // 图片起始列
	 * @param row1       // 图片起始行
	 * @throws IOException 异常
	 */
	public void addPicture(String imagePath, SheetModel sheetModel, int col1, int row1, Cell cell) throws IOException {
		// 获取工作簿实例
		Workbook workbook = sheetModel.getSheet().getWorkbook();
		InputStream inputStream = new FileInputStream(imagePath);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		inputStream.close();
		// 图片转换为BufferedImage对象
		BufferedImage outputImage = ImageIO.read(new ByteArrayInputStream(imageBytes));
		//图片原始宽高
		double originalWidth = outputImage.getWidth();
		double originalHeight = outputImage.getHeight();
		// 获取单元格宽度和高度
		double cellWidth = sheetModel.getSheet().getColumnWidthInPixels(cell.getColumnIndex());

		Drawing<?> drawing = sheetModel.getSheet().createDrawingPatriarch();
		ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col1 + 1, row1 + 1);
		anchor.setAnchorType(DONT_MOVE_AND_RESIZE);
		// 创建图片
		drawing.createPicture(anchor, workbook.addPicture(
				imageBytes, Workbook.PICTURE_TYPE_PNG));
        //设置图片大小，未设置成功
		// drawing.createPicture(anchor, pictureIdx).resize(scaleX, scaleY);
	}

	/**
	 * @param imagePathList  多条图片路径
	 * @param sheetModel 工作表
	 * @param col1       // 图片起始列
	 * @param row1       // 图片起始行
	 * @throws IOException 异常
	 */
	public void addPicture(List<String> imagePathList, SheetModel sheetModel, int col1, int row1, Cell cell) throws IOException {
		// 获取工作簿实例
		Drawing<?> drawing = sheetModel.getSheet().createDrawingPatriarch();
		Workbook workbook = sheetModel.getSheet().getWorkbook();
		ClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, col1, row1, col1 + 1, row1 + 1);
		anchor.setAnchorType(MOVE_AND_RESIZE);
		for (String  imagePath:imagePathList) {
			File imageFile = new File(imagePath);
			if (!imageFile.exists()) {
				System.out.println("警告：文件 " + imagePath + " 不存在，跳过此图片插入操作！");
				continue; // 跳过当前不存在的文件，处理下一个图片
			}
			InputStream inputStream = new FileInputStream(imagePath);
			byte[] imageBytes = IOUtils.toByteArray(inputStream);
			inputStream.close();
			BufferedImage outputImage = ImageIO.read(new ByteArrayInputStream(imageBytes));
			//图片原始宽高
			double originalWidth = outputImage.getWidth();
			double originalHeight = outputImage.getHeight();
//			sheetModel.getSheet().setColumnWidth(col1, (int) (originalWidth/100*255));
//			sheetModel.getSheet().getRow(row1).setHeight((short)(originalHeight/100*255));
						// 创建图片
			drawing.createPicture(anchor, workbook.addPicture(
					imageBytes, Workbook.PICTURE_TYPE_PNG));
		}

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
