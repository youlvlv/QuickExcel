package com.lizhiwei.quickExcel.util;

import com.lizhiwei.quickExcel.core.FastExcelCore;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.entity.WriteWorkType;
import com.lizhiwei.quickExcel.model.ExcelBaseModel;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;

import java.text.SimpleDateFormat;
import java.util.List;
import java.util.function.Consumer;

/**
 * 下载excel
 */
public class DownloadExcel extends ExcelBaseModel {

	private static final SimpleDateFormat df = new SimpleDateFormat("MM月dd日");

	private static final FastExcelCore FAST_EXCEL_CORE = new FastExcelCore();

	/**
	 * 生成EXCEL表
	 *
	 * @param operation   文件操作
	 * @param entity      列表实体类
	 * @param listContent 列表
	 * @param <T>         实体类
	 */
	public static <T> void setExcelProperty(FileOperation operation, Class<T> entity, List<T> listContent,
	                                        IndexType indexType, Consumer<CellStyle> styleConsumer) {
		//列表排序
		try {
			//创建表格工作空间
			ExcelModel excel = new ExcelModel();
			if (styleConsumer != null) {
				styleConsumer.accept(excel.getDefaultStyle());
			}
			//创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
			SheetModel sheet = excel.newSheet();
			if (indexType != IndexType.NULL) {
				sheet.createSerialNumber(indexType);
			}
			//set Sheet页头部
			sheet.createHeader(entity);
			//set Sheet页内容
			sheet.createContent(entity, listContent);
			excel.exportExcel(operation).close();
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("导出表格时出现异常...请联系管理员", e);
		}
	}

	public static <T> void setExcelProperty(FileOperation operation, Class<T> entity, List<T> listContent,
											IndexType indexType) {
		//列表排序
		try {
			//创建表格工作空间
			ExcelModel excel = new ExcelModel();
			//创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
			SheetModel sheet = excel.newSheet();
			if (indexType != IndexType.NULL) {
				sheet.createSerialNumber(indexType);
			}
			//set Sheet页头部
			sheet.createHeader(entity);
			//set Sheet页内容
			sheet.createContent(entity, listContent);
			excel.exportExcel(operation).close();
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("导出表格时出现异常...请联系管理员", e);
		}
	}


	public static <T> void setExcelProperty(FileOperation operation, List<ExcelEntity> entity, List<T> listContent,
	                                        IndexType indexType,Consumer<CellStyle> styleConsumer) {
		//列表排序
		try {
			//创建表格工作空间
			ExcelModel excel = new ExcelModel();
			if (styleConsumer != null) {
				styleConsumer.accept(excel.getDefaultStyle());
			}
			//创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
			SheetModel sheet = excel.newSheet();
			if (indexType != IndexType.NULL) {
				sheet.createSerialNumber(indexType);
			}
			//set Sheet页头部
			sheet.createHeader(entity);
			//set Sheet页内容
			sheet.createContent(entity, listContent);
			excel.exportExcel(operation).close();
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("导出表格时出现异常...请联系管理员", e);
		}
	}

	public static <T> void setExcelProperty(FileOperation operation, List<ExcelEntity> entity, List<T> listContent,
											IndexType indexType,List<String> imgList) {
		//列表排序
		try {
			//创建表格工作空间
			ExcelModel excel = new ExcelModel();
			//创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
			SheetModel sheet = excel.newSheet();
			if (indexType != IndexType.NULL) {
				sheet.createSerialNumber(indexType);
			}
			//set Sheet页头部
			sheet.createHeader(entity);
			//set Sheet页内容
			sheet.createContent(entity, listContent);
			excel.exportExcel(operation).close();
		} catch (Exception e) {
			e.printStackTrace();
			throw new RuntimeException("导出表格时出现异常...请联系管理员", e);
		}
	}

	/**
	 * 生成EXCEL表
	 *
	 * @param fileNameParam 文件名
	 * @param response      下载流
	 * @param entity        列表实体类
	 * @param listContent   列表
	 * @param <T>           实体类
	 */
    public static <T> void setExcelProperty(FileOperation operation, Class<T> entity, List<T> listContent) {
        setExcelProperty(operation, entity, listContent, IndexType.NULL, null);
    }



    /**
     * 生成EXCEL表
     *
     * @param operation   文件操作
     * @param response      下载流
     * @param entity        实体类
     * @param listContent   列表
     * @param styleConsumer 自定义样式
     * @param <T>
     */
    public static <T> void setExcelProperty(FileOperation operation, Class<T> entity,
                                            List<T> listContent, Consumer<CellStyle> styleConsumer) {
        setExcelProperty(operation, entity, listContent, IndexType.NULL, styleConsumer);
    }


	public static <T> void createExcel(FileOperation operation, Class<T> entity,
	                                   List<T> listContent) {
		setExcelProperty(operation, entity, listContent, IndexType.NULL);
	}

	public static <T> void createExcel(FileOperation operation, Class<T> entity,
	                                   List<T> listContent, WriteWorkType writeWorkType) {

		switch (writeWorkType) {
		case QuickExcel -> createExcel(operation, entity, listContent);
		case FastExcel -> FAST_EXCEL_CORE.createExcel(operation, entity, listContent);
		}
	}

	private DownloadExcel() {
	}
}
