package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.entity.ExcelImg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.*;

public class ReadExcelImg {

	private static final String OS = System.getProperty("os.name").toLowerCase();


	/**
	 * 从指定Excel文件中读取图片信息并将其保存到ExcelImg对象列表中.
	 *
	 * @param excelFilePath Excel文件的路径
	 * @return 包含图片位置信息和关联单元格值的ExcelImg对象列表
	 * @throws Exception 如果在处理Excel文件时发生任何异常，则抛出异常
	 */
	public static List<ExcelImg> importAndSaveImages(String excelFilePath, String filename) throws Exception {
		File target = new File(excelFilePath, filename);
		FileInputStream fi = new FileInputStream(target);

		try (XSSFWorkbook workbook = new XSSFWorkbook(fi)) {
			// 获取第一个工作表
			XSSFSheet sheet = workbook.getSheetAt(0);

			// 获取绘图 patriarch 对象
			XSSFDrawing drawing = sheet.getDrawingPatriarch();

			// 初始化存储图片信息的对象列表
			List<ExcelImg> list = new ArrayList<>();
			// 遍历所有形状（包括图片）
			for (XSSFShape shape : drawing.getShapes()) {
				// 判断当前形状是否为图片
				if (shape instanceof XSSFPicture) {
					XSSFPicture picture = (XSSFPicture) shape;
					// 获取图片在单元格中的位置信息
					XSSFClientAnchor ctMarker = picture.getPreferredSize();
					int row = ctMarker.getRow1();
					// 获取图片所在单元格的值
					String noNum = readCell(sheet, row);
					// 创建并初始化ExcelImg对象
					ExcelImg excelImg = new ExcelImg();
					excelImg.setNoNum(noNum);
					excelImg.setRow(String.valueOf(row));
					// 获取图片二进制数据
					byte[] imgData = getPictureData(picture);
					// 保存图片到指定路径
					String path = saveImage(imgData, getPictureExtension(picture));
					excelImg.setPath(path);
					// 将ExcelImg对象添加到列表中
					list.add(excelImg);
				}
			}
			return list;
		} catch (Exception e) {
			// 记录或处理异常
			e.printStackTrace();
			return new ArrayList<>();
		}
	}

	/**
	 * 读取指定单元格的字符串内容
	 *
	 * @param sheet 工作表对象
	 * @param row   行索引
	 * @return 单元格的字符串内容，如果单元格为空则返回null
	 */
	private static String readCell(Sheet sheet, int row) {
		Row dataRow = sheet.getRow(row);
		if (dataRow != null) {
			Cell cell = dataRow.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
			String stringValue;
			double numericValue = 0;
			switch (cell.getCellType()) {
				case STRING:
					stringValue = cell.getStringCellValue();
					// 处理字符串类型数据
					break;
				case NUMERIC:
					numericValue = cell.getNumericCellValue();
					// 如果数字应该表示日期或时间，请进一步处理（例如：Date date = DateUtil.getJavaDate(numericValue, false);）
					break;
				case BOOLEAN:
					boolean boolValue = cell.getBooleanCellValue();
					// 处理布尔类型数据
					break;
				case ERROR:
					// 处理错误类型单元格
					break;
				case FORMULA:
					// 公式类型单元格，可能需要根据实际情况获取计算结果
					switch (cell.getCachedFormulaResultType()) {
						case STRING:
							stringValue = cell.getStringCellValue();
							break;
						case NUMERIC:
							numericValue = cell.getNumericCellValue();
							break;
						// 类似地，添加其他类型处理...
					}
					break;
				default:
					throw new IllegalStateException("Unknown cell type encountered: " + cell.getCellType());
			}
			if (numericValue != 0) {
				return String.valueOf((int) numericValue);
			}
		}
		return null;
	}

	/**
	 * 获取图片二进制数据
	 *
	 * @param picture 图片对象
	 * @return 图片二进制数据
	 */
	private static byte[] getPictureData(XSSFPicture picture) {
		XSSFPictureData pictureData = picture.getPictureData();
		return pictureData.getData();
	}

	/**
	 * MIME类型和后缀对应关系
	 */
	private static final Map<String, String> MIME_TYPE_TO_EXTENSION = new HashMap<>() {{
		put("image/jpeg", ".jpg");
		put("image/png", ".png");
		put("image/gif", ".gif");
		put("image/bmp", ".bmp");
		put("image/x-emf", ".emf");
		put("image/unknown",".jpg");
		// 添加其他需要支持的MIME类型和后缀对应关系
	}};

	// 获取图片的MIME类型并转换为文件后缀
	public static String getPictureExtension(XSSFPicture picture) {
		String mimeType = picture.getPictureData().getMimeType();
		return MIME_TYPE_TO_EXTENSION.getOrDefault(mimeType, ".jpg");
	}

	/**
	 * 判断是否是Windows系统
	 *
	 * @return 返回
	 */
	public static boolean isWindows() {
		String os = System.getProperty("os.name").toLowerCase();
		return os.contains("win");
	}


	/**
	 * 文件存放真实路径
	 *
	 * @return 返回
	 */


	public static String getRealPath() {
		String realPath = "";
		if (OS.contains("linux")){
			realPath="/datadisk/hsoa/out/";
		} else {
			realPath ="D:\\datadisk\\hsoa\\out\\";
		}
		return realPath;
	}


	/**
	 * 保存图片到指定路径
	 *
	 * @param imgData 图片二进制数据
	 * @param type    图片类型
	 */
	private static String saveImage(byte[] imgData, String type) throws IOException {
		String imagePath = getRealPath();
		// 创建保存路径目录（如果不存在）
		File saveDir = new File(imagePath);
		if (!saveDir.exists()) {
			boolean created = saveDir.mkdirs();
			if (!created) {
				throw new RuntimeException("无法创建保存路径目录");
			}
		}
		// 创建日期子文件夹
		LocalDate currentDate = LocalDate.now();
		String dateFolderName = currentDate.toString().replace("-", "");
		File dateFolder = new File(imagePath, dateFolderName);
		if (!dateFolder.exists()) {
			boolean created = dateFolder.mkdirs();
			if (!created) {
				throw new RuntimeException("无法创建日期子文件夹");
			}
		}
		String randomFileName = UUID.randomUUID().toString().replace("-", "") + type;
		String filePath = dateFolder.getAbsolutePath() + File.separator + randomFileName;
		try (FileOutputStream fos = new FileOutputStream(filePath)) {
			fos.write(imgData);
		}
		return File.separator + dateFolderName + File.separator + randomFileName;
	}
}
