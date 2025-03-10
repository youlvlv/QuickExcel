package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * excel模型
 */
public class ExcelModel extends ExcelBaseModel implements AutoCloseable {

	/**
	 * 默认的单元格格式
	 */
	protected CellStyle DEFAULT_CELL_STYLE;

	protected Workbook xWorkbook;

	protected CellStyle CustomCellStyle;

	public ExcelModel() {
		xWorkbook = new XSSFWorkbook();
		DEFAULT_CELL_STYLE = xWorkbook.createCellStyle();
		//设置水平、垂直居中
		DEFAULT_CELL_STYLE.setAlignment(HorizontalAlignment.CENTER);
		DEFAULT_CELL_STYLE.setVerticalAlignment(VerticalAlignment.CENTER);
		//设置字体
		Font headerFont = xWorkbook.createFont();
		headerFont.setFontHeightInPoints((short) 12);
		/*headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);*/
		headerFont.setFontName("宋体");
		DEFAULT_CELL_STYLE.setFont(headerFont);
		DEFAULT_CELL_STYLE.setWrapText(true);//是否自动换行
	}

	public ExcelModel(Workbook workbook) {
		this.xWorkbook = workbook;
		DEFAULT_CELL_STYLE = xWorkbook.createCellStyle();
	}

	/**
	 * 获取poi模型
	 *
	 * @return
	 */
	public Workbook getWorkbook() {
		return xWorkbook;
	}

	/**
	 * 通过流写出
	 *
	 * @param stream
	 */
	public void write(OutputStream stream) throws IOException {
		xWorkbook.write(stream);
	}

	/**
	 * 创建新 sheet
	 * 并填充数据
	 *
	 * @param name        sheet 名
	 * @param entity      实体类
	 * @param listContent 内容
	 * @param <T>         实体类
	 * @return
	 */
	public <T> ExcelModel newSheet(String name, Class<T> entity, List<T> listContent) {
		SheetModel sheet = this.newSheet(name);
		List<ExcelEntity> list = util().getExcelEntities(entity);
		util().setSheetHeader(sheet, list);
		util().setSheetContent(sheet, listContent, list);
		return this;
	}

	/**
	 * 创建新 sheet
	 *
	 * @param name sheet 名
	 * @return
	 */
	public SheetModel newSheet(String name) {
		Sheet xSheet = xWorkbook.createSheet(name);
		return new SheetModel(xSheet, this);
	}

	/**
	 * 创建新 sheet
	 *
	 * @return
	 */
	public SheetModel newSheet() {
		Sheet xSheet = xWorkbook.createSheet();
		return new SheetModel(xSheet, this);
	}

	/**
	 * 打开指定名称的sheet
	 *
	 * @param name
	 * @return
	 */
	public SheetModel openSheet(String name) {
		Sheet xSheet = xWorkbook.getSheet(name);
		return new SheetModel(xSheet, this, true);
	}

	/**
	 * 打开指定索引的sheet
	 *
	 * @param index
	 * @return
	 */
	public SheetModel openSheetAt(int index) {
		Sheet xSheet = xWorkbook.getSheetAt(index);
		return new SheetModel(xSheet, this, true);
	}

//    public ExcelModel exportExcel(FileOutputStream stream) {
//        try {
//            xWorkbook.write(stream);
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }
//        return this;
//    }

	/**
	 * 获取默认单元格格式
	 *
	 * @return
	 */
	public CellStyle getDefaultStyle() {
		return DEFAULT_CELL_STYLE;
	}

	public ExcelModel setDefaultStyle() {
		return this;
	}

	/**
	 * 导出excel
	 *
	 * @param operation 文件操作类
	 * @return
	 */
	public ExcelModel exportExcel(FileOperation operation) {
		operation.run(this);
		return this;
	}

	/**
	 * 导出excel并关闭excel
	 *
	 * @param operation
	 */
	public void exportExcelAndClose(FileOperation operation) {
		operation.run(this);
		this.close();
	}

	public void close() {
		try {
			xWorkbook.close();
		} catch (IOException e) {
			throw new RuntimeException(e);
		} finally {

		}
	}

	/**
	 * 设置对角线
	 *
	 * @param startRow   开始行
	 * @param endRow     结束行
	 * @param sheetModel sheetModel
	 */
	public void diagonalLine(int startRow, int endRow, SheetModel sheetModel) {
		Drawing patriarch = sheetModel.createDrawingPatriarch();
		XSSFSimpleShape line = ((XSSFDrawing) patriarch).createSimpleShape(new XSSFClientAnchor(0, 0, 0, 100, (short) 0, startRow, (short) 0, endRow));
		line.setShapeType(ShapeTypes.LINE); // 设置形状类型为线条
		line.setLineWidth(1.0); // 设置线条宽度
		line.setLineStyleColor(0, 0, 0); // 设置线条颜色为黑色
		// 设置线条的起始和终止坐标，这里仅作示例，请根据实际需求调整
		int x1 = 1; // 左上角x坐标（EMU单位）
		int y1 = 50; // 左上角y坐标（EMU单位）
		int x2 = 150; // 右下角x坐标（EMU单位）
		int y2 = 100; // 右下角y坐标（EMU单位）
		line.getAnchor().setDx1(x1 * Units.EMU_PER_PIXEL);
		line.getAnchor().setDy1(y1 * Units.EMU_PER_PIXEL);
		line.getAnchor().setDx2(x2 * Units.EMU_PER_PIXEL);
		line.getAnchor().setDy2(y2 * Units.EMU_PER_PIXEL);
	}


	public CellStyle createStyle() {
		return xWorkbook.createCellStyle();
	}

	/**
	 * @param imagePath  图片路径
	 * @param sheetModel 工作表
	 * @param col1       // 图片起始列
	 * @param row1       // 图片起始行
	 * @param col2       // 图片结束列
	 * @param row2       // 图片结束行
	 * @throws IOException 异常
	 */
	public void addPicture(String imagePath, SheetModel sheetModel, int col1, int row1, int col2, int row2) throws IOException {
		InputStream inputStream = new FileInputStream(imagePath);
		byte[] imageBytes = IOUtils.toByteArray(inputStream);
		inputStream.close();
		int pictureIdx = xWorkbook.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
		CreationHelper helper = xWorkbook.getCreationHelper();
		Drawing drawing = sheetModel.createDrawingPatriarch();
		//添加位置
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setCol1(col1); // 图片起始列
		anchor.setRow1(row1); // 图片起始行
		anchor.setCol2(col2); // 图片结束列
		anchor.setRow2(row2); // 图片结束行
		// 创建图片并设置锚点
		Picture picture = drawing.createPicture(anchor, pictureIdx);
//        picture.resize();//自适应大小
	}


	public ExcelType getExcelType() {
		if (xWorkbook instanceof XSSFWorkbook || xWorkbook instanceof SXSSFWorkbook) {
			return ExcelType.XLSX;
		} else if (xWorkbook instanceof HSSFWorkbook) {
			return ExcelType.XLS;
		}
		throw new IORunTimeException("当前不支持该类型");
	}


	public enum ExcelType {


		XLS("xls"),
		XLSX("xlsx");

		private final String file_ext;

		ExcelType(String file_ext) {
			this.file_ext = file_ext;
		}

		public String getFile_ext() {
			return file_ext;
		}
	}
}

