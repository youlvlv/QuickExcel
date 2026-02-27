package com.lizhiwei.quickExcel.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.model.UploadFile;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Excel 读取器统一接口
 * V2 和 V3 引擎都实现此接口
 */
public interface ExcelReader {
    
    /**
     * 读取 Excel 文件
     */
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity);
    
    /**
     * 读取 Excel 文件（带 safe 模式）
     */
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe);
    
    /**
     * 读取 Excel 文件（带 safe 模式和 readImage 选项）
     */
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe, boolean readImage);
    
    /**
     * 读取 Excel 文件（使用指定属性列表）
     */
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, 
                          boolean safe, boolean readImage, List<ExcelEntity> propertieList);
    
    /**
     * 读取 Excel 为 Map 列表
     */
    List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, int sheetNum, 
                                              boolean safe, List<ExcelEntity> propertieList);
    
    /**
     * 从 UploadFile 读取 Excel
     */
    <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity);
    
    /**
     * 从 UploadFile 读取 Excel（带 safe 模式）
     */
    <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe);
    
    /**
     * 从 UploadFile 读取 Excel（带 safe 和 readImage 选项）
     */
    <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity, 
                          boolean safe, boolean readImage);
    
    /**
     * 按 Sheet 名称读取 Excel
     */
    <T> List<T> readExcel(File file, int startRow, int startCol, String sheetName, Class<T> entity, 
                          boolean safe, boolean readImage);
}
