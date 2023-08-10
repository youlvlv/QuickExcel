package com.lizhiwei.quickExcel.util;

import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.ExcelBaseModel;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.SheetModel;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.List;

/**
 * 下载excel
 */
public class DownloadExcel extends ExcelBaseModel {

	/**
	 * 生成EXCEL表
	 *
	 * @param operation   文件操作
	 * @param entity      列表实体类
	 * @param listContent 列表
	 * @param <T>         实体类
	 */
	public static <T> void setExcelProperty(FileOperation operation, Class<T> entity, List<T> listContent, IndexType indexType) {
		SimpleDateFormat df = new SimpleDateFormat("MM月dd日");
		//列表排序
		try {
			//创建表格工作空间
			ExcelModel excel = new ExcelModel();
			//创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
			SheetModel sheet = excel.newSheet();
			if (indexType != IndexType.NULL){
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
	public static <T> void setExcelProperty(String fileNameParam, HttpServletResponse response, Class<T> entity, List<T> listContent) {
		setExcelProperty(new DefaultDownloadExcel(response, fileNameParam), entity, listContent, IndexType.NULL);
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
	public static <T> void setExcelProperty(String fileNameParam, HttpServletResponse response, Class<T> entity, List<T> listContent, IndexType type) {
		setExcelProperty(new DefaultDownloadExcel(response, fileNameParam), entity, listContent, type);
	}

	private DownloadExcel() {
	}
}
