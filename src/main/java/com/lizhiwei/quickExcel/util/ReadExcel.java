package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.PictureMap;
import com.lizhiwei.quickExcel.entity.ReadErrorInfo;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;
import com.lizhiwei.quickExcel.model.ExcelBaseModel;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.UploadFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFDrawing;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class ReadExcel extends ExcelBaseModel {

	public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity) {
		return readExcel(file, startrow, startcol, sheetnum, entity, false, false);
	}

	public static ExcelModel readExcel(File file) {
		try {
			return new ExcelModel(getWorkbook(file));
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	public static <T> List<T> readExcel(File file, int startrow, int startcol, String sheetName, Class<T> entity,
	                                    boolean safe, boolean readImage) {
		int sheetnum = 0;
		try {
			Workbook wb = getWorkbook(file);
			sheetnum = wb.getSheetIndex(wb.getSheet(sheetName));
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		return readExcel(file, startrow, startcol, sheetnum, entity, safe, readImage);
	}

	public static List<Map<String, String>> readExcel(File file, int startrow, int startcol, int sheetnum, boolean safe, List<ExcelEntity> propertieList) {
		List<Map<String, String>> varList = new ArrayList<>();
		boolean error = false;
		List<ReadErrorInfo> errorInfoList = new ArrayList<>();
		try {
			Workbook wb = getWorkbook(file);
			Sheet sheet = wb.getSheetAt(sheetnum); // sheet 从0开始
			List<ExcelEntity> properties = getExcelEntities(startrow, startcol, propertieList, sheet);
			Row row;
			//循环实体类所有属性
			int rowNum = sheet.getLastRowNum() + 1; // 取得最后一行的行号
			//空行数
			int emptySize = 0;
			/*--------------数据行-----------------------*/
			for (int i = startrow; i < rowNum; i++) { // 行循环开始

				row = sheet.getRow(i); // 行
				if (row == null) {
					break;
				}
				Map<String, String> t = new HashMap<>();
				//获取需要读取的数量
				int size = properties.size();
				for (ExcelEntity property : properties) {
					// 查看该字段是否允许导入
					Field field = null;
					Method method = null;
					try {
						//读取当前字段在excel中的值
						String o = getExcelStringValue(wb, getMergedRegionValue(sheet, i, property.getValue()), property);
						//若当前字段为空，则读取数量减1
						if (o == null || o.isEmpty()) {
							--size;
						}
						t.put(property.getProperty(), o);
					} catch (ExcelValueException e) {
						if (safe) {
							error = true;
							errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
						} else {
							throw new ExcelReadException("第" + i + "行" + " " + e.getMessage(), e);
						}
					}

				}
				//若当前行为空行则将连续空行+1
				if (size == 0) {
					//连续三行都是空行，则认定当前为excel结尾
					if (++emptySize > 3) {
						break;
					}
				} else {
					//若不为空行，则清空连续空行数
					emptySize = 0;
					varList.add(t);
				}
			}
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		if (error) {
			throw new ExcelReadException(errorInfoList);
		}
		if (varList.isEmpty()) {
			throw new ExcelReadException("当前表格为空");
		}
		return varList;

	}

	public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity,
	                                    boolean safe, boolean readImage, List<ExcelEntity> propertieList) {
		List<T> varList = new ArrayList<>();
		boolean error = false;
		List<ReadErrorInfo> errorInfoList = new ArrayList<>();
		try {
			Workbook wb = getWorkbook(file);
			Sheet sheet = wb.getSheetAt(sheetnum); // sheet 从0开始
			PictureMap pictureMap = new PictureMap();
			// 获取绘图 patriarch 对象
			if (readImage) {
//				if (wb instanceof XSSFWorkbook xssfWorkbook) {
//					// 获取第一个工作表
//					XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
//
//					// 获取绘图 patriarch 对象
//					XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
//				}
				Drawing<?> drawing = sheet.getDrawingPatriarch();
				Optional.ofNullable(drawing).ifPresent(draw -> {
					for (Object o : draw) {
						if (o instanceof Picture picture) {
							ClientAnchor ca = picture.getClientAnchor();
							pictureMap.put(ca.getRow1(), picture);
						}
					}
				});

			}


			List<ExcelEntity> properties = getExcelEntities(startrow, startcol, propertieList, sheet);
			Row row;
			//循环实体类所有属性
			int rowNum = sheet.getLastRowNum() + 1; // 取得最后一行的行号
			//空行数
			int emptySize = 0;
			/*--------------数据行-----------------------*/
			for (int i = startrow; i < rowNum; i++) { // 行循环开始

				row = sheet.getRow(i); // 行
				if (row == null) {
					break;
				}

				T t = null;
				try {
					//创建新的实体类
					t = entity.getDeclaredConstructor().newInstance();
				} catch (InstantiationException | IllegalAccessException | InvocationTargetException |
				         NoSuchMethodException e) {
					throw new RuntimeException("构建实体类失败！请检查实体类", e);
				}
				int pictureIndex = 0;
				//获取需要读取的数量
				int size = properties.size();
				for (ExcelEntity property : properties) {
					// 查看该字段是否允许导入
					Field field = null;
					Method method = null;
					try {
						//读取当前字段在excel中的值
						Cell cell = getMergedRegionValue(sheet, i, property.getValue());
						Object o = getExcelValue(wb, cell, property);
						//若当前字段为空，则读取数量减1
						if (o == null || o.toString().isEmpty()) {
							--size;
						}
						switch (property.getParamType()) {
							//若为属性
							case FIELD: {
								//实例化字段
								field = entity.getDeclaredField(property.getProperty());
								field.setAccessible(true);
								//赋值
								field.set(t, o);
								break;
							}
							//若为方法
							case METHOD: {
								String set = "set" + Pattern.compile("^.").matcher(property.getProperty()).replaceFirst(m -> m.group().toUpperCase());
								method = entity.getMethod(set, property.getType());
								//赋值
								method.invoke(t, o);
								break;
							}

							case IMAGE: {
								Picture picture = pictureMap.get(row.getRowNum(), pictureIndex++);
								field = entity.getDeclaredField(property.getProperty());
								field.setAccessible(true);
								//赋值
								field.set(t, formatValue(property,
										ExcelConfig.getImageFileFunction().apply(new ByteArrayInputStream(picture.getPictureData().getData()), getPictureExtension(picture))));
								break;
							}

						}


					} catch (NoSuchFieldException | IllegalAccessException | NoSuchMethodException |
					         InvocationTargetException e) {
						throw new RuntimeException(e);
					} catch (ExcelValueException e) {
						if (safe) {
							error = true;
							errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
						} else {
							throw new ExcelReadException("第" + i + "行" + " " + e.getMessage(), e);
						}
					}

				}
				//若当前行为空行则将连续空行+1
				if (size == 0) {
					//连续三行都是空行，则认定当前为excel结尾
					if (++emptySize > 3) {
						break;
					}
				} else {
					//若不为空行，则清空连续空行数
					emptySize = 0;
					varList.add(t);
				}
			}
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
		if (error) {
			throw new ExcelReadException(errorInfoList);
		}
		if (varList.isEmpty()) {
			throw new ExcelReadException("当前表格为空");
		}
		return varList;
	}


	private static Workbook getWorkbook(File file) throws IOException {
		//读取文件
		FileInputStream fi = new FileInputStream(file);
		String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
		Workbook wb = null;
		//判断文件类型
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(fi);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(fi);
		} else {
			throw new IORunTimeException("您导入的文件不是标准excel文件");
		}
		return wb;
	}

	/**
	 * 匹配头生成 ExcelEntity
	 *
	 * @param startrow
	 * @param startcol
	 * @param propertieList
	 * @param sheet
	 * @return
	 */
	private static List<ExcelEntity> getExcelEntities(int startrow, int startcol, List<ExcelEntity> propertieList, Sheet sheet) {
		List<ExcelEntity> properties = new ArrayList<>();

		/*----------匹配头------------*/
		Row row = sheet.getRow(startrow - 1); // 行
		int cellNum = row.getLastCellNum(); // 每行的最后一个单元格位置
		//首行名称与位置
		Map<String, Integer> cellName = new HashMap<>();
		for (int j = startcol; j < cellNum; j++) { // 列循环开始
			cellName.put(getCellValue(getMergedRegionValue(sheet, startrow - 1, j)), j);
		}
		for (ExcelEntity excelEntity : propertieList) {
			if ((cellName.containsKey(excelEntity.getTitle()) || (excelEntity.getAlias().isEmpty() && cellName.containsKey(excelEntity.getAlias()))) && excelEntity.isRead()) {
				excelEntity.setValue(cellName.get(excelEntity.getTitle()));
				//实体类中该属性类型
				properties.add(excelEntity);
			}
		}
		return properties;
	}

	/**
	 * 读取excel信息
	 * 默认0
	 *
	 * @param file     excel文件
	 * @param startrow 开始行
	 * @param startcol 开始列
	 * @param sheetnum sheet号
	 * @param entity   实体类
	 * @param safe     是否综合报错
	 * @param <T>
	 * @return 列表信息
	 */
	public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity,
	                                    boolean safe, boolean readImage) {
		return readExcel(file, startrow, startcol, sheetnum, entity, safe, readImage, getExcelEntities(entity));
	}


	/**
	 * 读取excel
	 *
	 * @param filepath 文件路径
	 * @param filename 文件名
	 * @param startrow 开始行号
	 * @param startcol 开始列号
	 * @param sheetnum sheet
	 * @return list
	 */
	public static <T> List<T> readExcel(String filepath, String filename, int startrow, int startcol, int sheetnum, Class<T> entity) {
		File target = new File(filepath, filename);
		return readExcel(target, startrow, startcol, sheetnum, entity);
	}

	/**
	 * 读取excel
	 *
	 * @param file     上传文件
	 * @param startrow 开始行号
	 * @param startcol 开始列号
	 * @param sheetnum sheet
	 * @return list
	 */
	public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity) {
		return readExcel(file.getFile(), startrow, startcol, sheetnum, entity);
	}

	public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity,
	                                    boolean safe, boolean readImage) {
		return readExcel(file.getFile(), startrow, startcol, sheetnum,entity, safe,readImage);
	}

	private static String getExcelStringValue(Workbook workbook, Cell cell, ExcelEntity property) {
		String cellValue = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		if (null != cell) {
			cellValue = getCellValue(workbook, cell, cellValue, sdf, property.getAccuracy());
			// 判断当前字段是否允许非空，并判断非空
			if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
				throw new ExcelValueException(property.getTitle() + "为空");
			} else if (!cellValue.trim().isEmpty()) {
				return formatValue(property, cellValue);
			}
		}
		return null;
	}

	private static String formatValue(ExcelEntity property, String cellValue) {
		if (property.getFormat() != null) {
			ExcelFormatBase<?> format = property.getFormat();
			try {
				if (format instanceof DefaultFormat) {
					return ((DefaultFormat) format).ReadToExcel(String.class, cellValue).toString();
				}
				return format.ReadToExcel(cellValue).toString();
			} catch (Exception e) {
				throw new ExcelValueException(property.getTitle() + "错误", e);
			}
		}
		return cellValue;
	}

	/**
	 * 获取单元格值
	 *
	 * @param cell     单元格
	 * @param property 值类型
	 * @return 值
	 */
	private static Object getExcelValue(Workbook workbook, Cell cell, ExcelEntity property) {
		String cellValue = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		if (null != cell) {
			cellValue = getCellValue(workbook, cell, cellValue, sdf, property.getAccuracy());
			// 判断当前字段是否允许非空，并判断非空
			if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
				throw new ExcelValueException(property.getTitle() + "为空");
			} else if (!cellValue.trim().isEmpty()) {
				Class<?> type = property.getType();
				ExcelFormatBase<?> format = property.getFormat();
				try {
					if (format instanceof DefaultFormat) {
						return ((DefaultFormat) format).ReadToExcel(type, cellValue);
					}
					return format.ReadToExcel(cellValue);
				} catch (Exception e) {
					throw new ExcelValueException(property.getTitle() + "错误", e);
				}
			} else {
				return null;
			}
		}
		return null;
	}

	private static String getCellValue(Workbook workbook, Cell cell, String cellValue, SimpleDateFormat sdf, int accuracy) {
		switch (cell.getCellType()) { // 判断excel单元格内容的格式，并对其进行转换，以便插入数据库
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					//判断是否为日期类型
					cellValue = sdf.format(cell.getDateCellValue());
				} else {
					String msg = String.valueOf(cell.getNumericCellValue());
					if (msg.contains(".0")) {
						cellValue = checkNumber(String.valueOf(cell.getNumericCellValue()));
					} else {
						if (accuracy == -1) {
							cellValue = String.valueOf(cell.getNumericCellValue());
						} else {
							cellValue = BigDecimal.valueOf(cell.getNumericCellValue()).setScale(accuracy, RoundingMode.HALF_UP).toString();
						}
					}
				}
				break;
			case STRING:
				cellValue = cell.getStringCellValue();
				break;
			case BLANK:
				cellValue = "";
				break;
			case BOOLEAN:
				cellValue = String.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA:
				FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
				CellValue evaluate = formulaEvaluator.evaluate(cell);
				switch (evaluate.getCellType()) {
					case NUMERIC:
						cellValue = String.valueOf(evaluate.getNumberValue());
						break;
					case STRING:
						cellValue = evaluate.getStringValue();
						break;
				}
				break;
			case ERROR:
				cellValue = String.valueOf(cell.getErrorCellValue());
				break;
		}
		return cellValue;
	}


	/**
	 * 获取合并单元格的值
	 *
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static Cell getMergedRegionValue(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					return fRow.getCell(firstColumn);
				}
			}
		}

		return sheet.getRow(row).getCell(column);
	}

	/**
	 * 获取单元格的值
	 *
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		return cell.toString();
	}


	// 获取图片的MIME类型并转换为文件后缀
	public static String getPictureExtension(Picture picture) {
		String mimeType = picture.getPictureData().getMimeType();
		return MIME_TYPE_TO_EXTENSION.getOrDefault(mimeType, "");
	}

	// 工资条问题  上面的是原版的
	public static String checkNumber(String number) {

		String a = null;
		if (number.contains(".01") || number.contains(".02") || number.contains(".03") || number.contains(".04") || number.contains(".05")
				|| number.contains(".06") || number.contains(".07") || number.contains(".08") || number.contains(".09")) {
			a = number;
		} else {
			if (number.contains(".0")) {
				a = number.substring(0, number.length() - 2);
			} else if (number.contains("-0")) {
				a = number;
			} else {
				a = number;
			}
		}
		return a;
	}

}
