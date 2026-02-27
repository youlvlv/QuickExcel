package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.core.ExcelReader;
import com.lizhiwei.quickExcel.core.ExcelReaderFactory;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.UploadFile;
import com.lizhiwei.quickExcel.v2.ReadExcelByPoi;
import com.lizhiwei.quickExcel.v3.read.ReadExcelByXml;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Excel 读取工具类
 * <p>
 * 根据配置自动选择 V2（Apache POI）或 V3（SAX/DOM）解析引擎。
 * 默认使用 V2 引擎。
 * </p>
 * <p>
 * 配置方式：
 * <pre>
 * // 全局设置为 V3 引擎
 * ExcelConfig.setDefaultReadEngine(ExcelConfig.ReadEngine.V3);
 * </pre>
 * </p>
 */
public class ReadExcel {
    
    /**
     * 获取当前使用的读取器
     */
    private static ExcelReader getReader() {
        return ExcelReaderFactory.getReader();
    }
    
    // ==================== 基础读取方法 ====================
    
    /**
     * 读取 Excel 文件
     * @param file Excel 文件
     * @param startRow 开始行（从 0 开始）
     * @param startCol 开始列（从 0 开始）
     * @param sheetNum Sheet 索引（从 0 开始）
     * @param entity 实体类
     * @return 实体列表
     */
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    /**
     * 读取 Excel 文件（带 safe 模式）
     * @param safe 是否安全模式（收集所有错误后统一抛出）
     */
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                        boolean safe) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    /**
     * 读取 Excel 文件（带 safe 和 readImage 选项）
     * @param readImage 是否读取图片
     */
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                        boolean safe, boolean readImage) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage);
    }
    
    /**
     * 读取 Excel 文件（使用指定属性列表）
     */
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                        boolean safe, boolean readImage, List<ExcelEntity> propertieList) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage, propertieList);
    }
    
    /**
     * 读取 Excel 为 Map 列表
     */
    public static List<Map<String, String>> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                                      boolean safe, List<ExcelEntity> propertieList) {
        return getReader().readExcelAsMap(file, startRow, startCol, sheetNum, safe, propertieList);
    }
    
    // ==================== 按 Sheet 名称读取 ====================
    
    /**
     * 按 Sheet 名称读取 Excel
     */
    public static <T> List<T> readExcel(File file, int startRow, int startCol, String sheetName, Class<T> entity,
                                        boolean safe, boolean readImage) {
        return getReader().readExcel(file, startRow, startCol, sheetName, entity, safe, readImage);
    }
    
    public static <T> List<T> readExcel(File file, int startRow, int startCol, String sheetName, Class<T> entity,
                                        boolean safe) {
        int sheetNum = 0;
        try {
            Workbook wb = getWorkbook(file);
            sheetNum = wb.getSheetIndex(wb.getSheet(sheetName));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return readExcel(file, startRow, startCol, sheetNum, entity, safe, false);
    }
    
    // ==================== UploadFile 读取 ====================
    
    /**
     * 从 UploadFile 读取 Excel
     */
    public static <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    public static <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                        boolean safe) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    public static <T> List<T> readExcel(UploadFile file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                        boolean safe, boolean readImage) {
        return getReader().readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage);
    }
    
    // ==================== 路径读取 ====================
    
    /**
     * 从路径读取 Excel
     */
    public static <T> List<T> readExcel(String filepath, String filename, int startRow, int startCol, 
                                        int sheetNum, Class<T> entity) {
        File target = new File(filepath, filename);
        return readExcel(target, startRow, startCol, sheetNum, entity);
    }
    
    // ==================== 原始 Workbook 读取 ====================
    
    /**
     * 读取 Excel 为 ExcelModel（直接使用 POI）
     */
    public static ExcelModel readExcel(File file) {
        try {
            return new ExcelModel(getWorkbook(file));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    
    private static Workbook getWorkbook(File file) throws IOException {
        FileInputStream fi = new FileInputStream(file);
        String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
        Workbook wb = null;
        if (fileType.equals("xls")) {
            wb = new HSSFWorkbook(fi);
        } else if (fileType.equals("xlsx")) {
            wb = new XSSFWorkbook(fi);
        }
        return wb;
    }
    
    // ==================== 获取 ExcelEntity 列表 ====================
    
    /**
     * 获取实体类的 ExcelEntity 列表
     */
    public static <T> List<ExcelEntity> getExcelEntities(Class<T> entity) {
        return com.lizhiwei.quickExcel.model.ExcelBaseModel.getExcelEntities(entity);
    }
    
    // ==================== V2 引擎专用方法（直接调用）====================
    
    /**
     * 使用 V2 引擎读取 Excel
     */
    public static <T> List<T> readExcelByV2(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return ExcelReaderFactory.getV2Reader().readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    public static <T> List<T> readExcelByV2(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                             boolean safe) {
        return ExcelReaderFactory.getV2Reader().readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    public static <T> List<T> readExcelByV2(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                             boolean safe, boolean readImage) {
        return ExcelReaderFactory.getV2Reader().readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage);
    }
    
    // ==================== V3 引擎专用方法（直接调用）====================
    
    /**
     * 使用 V3 引擎读取 Excel
     */
    public static <T> List<T> readExcelByV3(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    public static <T> List<T> readExcelByV3(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                             boolean safe) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    public static <T> List<T> readExcelByV3(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                             boolean safe, ReadExcelByXml.ReadStrategy strategy) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe, strategy);
    }
    
    /**
     * 使用 V3 引擎按 sheet 名称读取 Excel
     */
    public static <T> List<T> readExcelByV3(File file, int startRow, int startCol, String sheetName, Class<T> entity) {
        return ReadExcelByXml.readExcelByName(file, startRow, startCol, sheetName, entity);
    }
    
    public static <T> List<T> readExcelByV3(File file, int startRow, int startCol, String sheetName, Class<T> entity,
                                             boolean safe) {
        return ReadExcelByXml.readExcelByName(file, startRow, startCol, sheetName, entity, safe);
    }
    
    public static List<Map<String, String>> readExcelByV3AsMap(File file, int startRow, int startCol, int sheetNum) {
        return ReadExcelByXml.readExcelAsMap(file, startRow, startCol, sheetNum, null);
    }
    
    public static List<Map<String, String>> readExcelByV3AsMap(File file, int startRow, int startCol, int sheetNum,
                                                                 List<String> headers) {
        return ReadExcelByXml.readExcelAsMap(file, startRow, startCol, sheetNum, headers);
    }
    
    /**
     * 使用 V3 引擎按 sheet 名称读取为 Map
     */
    public static List<Map<String, String>> readExcelByV3AsMap(File file, int startRow, int startCol, 
                                                                 String sheetName, List<String> headers) {
        return ReadExcelByXml.readExcelAsMapByName(file, startRow, startCol, sheetName, headers);
    }
    
    public static List<com.lizhiwei.quickExcel.v3.read.model.ImageData> readExcelImagesByV3(File file, int sheetNum) {
        return ReadExcelByXml.readExcelImages(file, sheetNum);
    }
    
    /**
     * 使用 V3 引擎按 sheet 名称读取图片
     */
    public static List<com.lizhiwei.quickExcel.v3.read.model.ImageData> readExcelImagesByV3(File file, String sheetName) {
        return ReadExcelByXml.readExcelImagesByName(file, sheetName);
    }
    
    // ==================== 兼容旧版方法（保留但标记为过时）====================
    
    /**
     * 使用 POI 直接读取（兼容旧版）
     * @deprecated 请使用 {@link #readExcelByV2(File, int, int, int, Class)}
     */
    @Deprecated
    public static <T> List<T> readExcelByPoi(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return ReadExcelByPoi.readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    /**
     * 使用 POI 直接读取（兼容旧版）
     * @deprecated 请使用 {@link #readExcelByV2(File, int, int, int, Class, boolean)}
     */
    @Deprecated
    public static <T> List<T> readExcelByPoi(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                              boolean safe) {
        return ReadExcelByPoi.readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    /**
     * 使用 XML 解析读取（兼容旧版）
     * @deprecated 请使用 {@link #readExcelByV3(File, int, int, int, Class)}
     */
    @Deprecated
    public static <T> List<T> readExcelByXml(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity);
    }
    
    /**
     * 使用 XML 解析读取（兼容旧版）
     * @deprecated 请使用 {@link #readExcelByV3(File, int, int, int, Class, boolean)}
     */
    @Deprecated
    public static <T> List<T> readExcelByXml(File file, int startRow, int startCol, int sheetNum, Class<T> entity,
                                              boolean safe) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
}
