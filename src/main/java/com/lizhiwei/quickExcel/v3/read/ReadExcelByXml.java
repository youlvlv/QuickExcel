package com.lizhiwei.quickExcel.v3.read;

import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;
import com.lizhiwei.quickExcel.v3.read.model.ExcelReadResult;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import com.lizhiwei.quickExcel.v3.read.parser.ImageParser;
import com.lizhiwei.quickExcel.v3.read.reader.DomExcelReader;
import com.lizhiwei.quickExcel.v3.read.reader.ExcelReader;
import com.lizhiwei.quickExcel.v3.read.reader.SaxExcelReader;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * V3 版本 Excel 读取工具类
 * <p>
 * 基于 XML 解析的高性能读取器，支持：
 * 1. SAX/DOM 两种解析模式
 * 2. 上下文缓存，避免重复解析文件元数据
 * 3. 一步读取实体类和图片数据
 * 4. 支持 DISPIMG 函数图片（WPS 嵌入单元格图片）
 * </p>
 */
public class ReadExcelByXml {
    
    private static final ExcelReader SAX_READER = new SaxExcelReader();
    private static final ExcelReader DOM_READER = new DomExcelReader();
    
    private static ExcelReader getReader(ReadStrategy strategy) {
        return switch (strategy) {
            case SAX -> SAX_READER;
            case DOM -> DOM_READER;
            default -> SAX_READER;
        };
    }
    
    // ==================== 上下文管理 ====================
    
    /**
     * 获取文件上下文（复用已解析的元数据）
     * @param file Excel 文件
     * @return 文件上下文
     */
    public static ExcelFileContext getContext(File file) {
        return ExcelFileContextManager.getInstance().getContext(file);
    }
    
    /**
     * 清理文件上下文缓存
     * @param file Excel 文件
     */
    public static void clearContext(File file) {
        ExcelFileContextManager.getInstance().removeContext(file);
    }
    
    /**
     * 清理所有上下文缓存
     */
    public static void clearAllContexts() {
        ExcelFileContextManager.getInstance().clearAll();
    }
    
    // ==================== Sheet 信息 ====================
    
    /**
     * 获取 sheet 名称对应的索引
     * @param file Excel 文件
     * @param sheetName sheet 名称
     * @return sheet 索引，找不到返回 -1
     */
    public static int getSheetIndex(File file, String sheetName) {
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        try {
            return context.getSheetIndex(sheetName);
        } catch (Exception e) {
            return -1;
        }
    }
    
    /**
     * 获取指定索引的 sheet 名称
     * @param file Excel 文件
     * @param sheetNum sheet 索引
     * @return sheet 名称，找不到返回 null
     */
    public static String getSheetName(File file, int sheetNum) {
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        try {
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetNum);
            return sheetInfo != null ? sheetInfo.name : null;
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 获取所有 sheet 名称
     * @param file Excel 文件
     * @return sheet 名称列表
     */
    public static List<String> getSheetNames(File file) {
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        try {
            List<String> names = new java.util.ArrayList<>();
            for (ExcelFileContext.SheetMetadata sheet : context.getSheetsMetadata()) {
                names.add(sheet.name);
            }
            return names;
        } catch (Exception e) {
            return new java.util.ArrayList<>();
        }
    }
    
    // ==================== 按索引读取实体类 ====================
    
    public static <T> List<T> readExcel(File file, int startRow, int sheetNum, Class<T> entity) {
        return readExcel(file, startRow, 0, sheetNum, entity, false, ReadStrategy.SAX);
    }
    
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return readExcel(file, startRow, startCol, sheetNum, entity, false, ReadStrategy.SAX);
    }
    
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe) {
        return readExcel(file, startRow, startCol, sheetNum, entity, safe, ReadStrategy.SAX);
    }
    
    public static <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                         Class<T> entity, boolean safe, ReadStrategy strategy) {
        ExcelReader reader = getReader(strategy);
        return reader.readExcel(file, startRow, startCol, sheetNum, entity, safe);
    }
    
    // ==================== 按名称读取实体类 ====================
    
    /**
     * 按 sheet 名称读取 Excel
     */
    public static <T> List<T> readExcelByName(File file, int startRow, String sheetName, Class<T> entity) {
        return readExcelByName(file, startRow, 0, sheetName, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称读取 Excel
     */
    public static <T> List<T> readExcelByName(File file, int startRow, int startCol, String sheetName, Class<T> entity) {
        return readExcelByName(file, startRow, startCol, sheetName, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称读取 Excel
     */
    public static <T> List<T> readExcelByName(File file, int startRow, int startCol, String sheetName, 
                                               Class<T> entity, boolean safe) {
        return readExcelByName(file, startRow, startCol, sheetName, entity, safe, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称读取 Excel
     */
    public static <T> List<T> readExcelByName(File file, int startRow, int startCol, String sheetName, 
                                               Class<T> entity, boolean safe, ReadStrategy strategy) {
        int sheetIndex = getSheetIndex(file, sheetName);
        if (sheetIndex < 0) {
            throw new ExcelReadException("找不到 Sheet: " + sheetName);
        }
        return readExcel(file, startRow, startCol, sheetIndex, entity, safe, strategy);
    }
    
    // ==================== 一步读取实体类和图片（新功能）====================
    
    /**
     * 一步读取 Excel 实体类和图片（包含浮动图片和 DISPIMG 图片）
     * <p>
     * 此方法会同时解析实体类数据和图片数据，并将图片自动绑定到标记为 @Excel(isPicture = true) 的字段
     * </p>
     * 
     * @param file Excel 文件
     * @param startRow 起始行（从 1 开始）
     * @param sheetNum sheet 索引（从 0 开始）
     * @param entity 实体类
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    public static <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int sheetNum, Class<T> entity) {
        return readExcelWithImages(file, startRow, 0, sheetNum, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 一步读取 Excel 实体类和图片
     * 
     * @param file Excel 文件
     * @param startRow 起始行（从 1 开始）
     * @param startCol 起始列（从 0 开始）
     * @param sheetNum sheet 索引（从 0 开始）
     * @param entity 实体类
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    public static <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int startCol, 
                                                              int sheetNum, Class<T> entity) {
        return readExcelWithImages(file, startRow, startCol, sheetNum, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 一步读取 Excel 实体类和图片
     * 
     * @param file Excel 文件
     * @param startRow 起始行（从 1 开始）
     * @param startCol 起始列（从 0 开始）
     * @param sheetNum sheet 索引（从 0 开始）
     * @param entity 实体类
     * @param safe 是否安全模式（出错不中断）
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    public static <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int startCol, 
                                                              int sheetNum, Class<T> entity, boolean safe) {
        return readExcelWithImages(file, startRow, startCol, sheetNum, entity, safe, ReadStrategy.SAX);
    }
    
    /**
     * 一步读取 Excel 实体类和图片（完整参数）
     * 
     * @param file Excel 文件
     * @param startRow 起始行（从 1 开始）
     * @param startCol 起始列（从 0 开始）
     * @param sheetNum sheet 索引（从 0 开始）
     * @param entity 实体类
     * @param safe 是否安全模式
     * @param strategy 读取策略（SAX/DOM）
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    public static <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int startCol, 
                                                              int sheetNum, Class<T> entity, 
                                                              boolean safe, ReadStrategy strategy) {
        ExcelReader reader = getReader(strategy);
        if (reader instanceof com.lizhiwei.quickExcel.v3.read.reader.AbstractExcelReader) {
            return ((com.lizhiwei.quickExcel.v3.read.reader.AbstractExcelReader) reader)
                    .readExcelWithImages(file, startRow, startCol, sheetNum, entity, safe);
        }
        // 降级处理：分别读取实体类和图片
        ExcelReadResult<T> result = new ExcelReadResult<>();
        List<T> entities = reader.readExcel(file, startRow, startCol, sheetNum, entity, safe);
        List<ImageData> images = reader.readExcelImages(file, sheetNum);
        result.setEntities(entities);
        result.addImages(images);
        return result;
    }
    
    /**
     * 按 sheet 名称一步读取 Excel 实体类和图片
     */
    public static <T> ExcelReadResult<T> readExcelWithImagesByName(File file, int startRow, 
                                                                    String sheetName, Class<T> entity) {
        return readExcelWithImagesByName(file, startRow, 0, sheetName, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称一步读取 Excel 实体类和图片
     */
    public static <T> ExcelReadResult<T> readExcelWithImagesByName(File file, int startRow, int startCol,
                                                                    String sheetName, Class<T> entity) {
        return readExcelWithImagesByName(file, startRow, startCol, sheetName, entity, false, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称一步读取 Excel 实体类和图片（完整参数）
     */
    public static <T> ExcelReadResult<T> readExcelWithImagesByName(File file, int startRow, int startCol,
                                                                    String sheetName, Class<T> entity, 
                                                                    boolean safe, ReadStrategy strategy) {
        int sheetIndex = getSheetIndex(file, sheetName);
        if (sheetIndex < 0) {
            throw new ExcelReadException("找不到 Sheet: " + sheetName);
        }
        return readExcelWithImages(file, startRow, startCol, sheetIndex, entity, safe, strategy);
    }
    
    // ==================== 读取为 Map ====================
    
    public static List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, 
                                                            int sheetNum, List<String> headers) {
        return readExcelAsMap(file, startRow, startCol, sheetNum, headers, ReadStrategy.SAX);
    }
    
    public static List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, 
                                                            int sheetNum, List<String> headers, ReadStrategy strategy) {
        ExcelReader reader = getReader(strategy);
        return reader.readExcelAsMap(file, startRow, startCol, sheetNum, headers);
    }
    
    /**
     * 按 sheet 名称读取 Excel 为 Map
     */
    public static List<Map<String, String>> readExcelAsMapByName(File file, int startRow, int startCol, 
                                                                  String sheetName, List<String> headers) {
        return readExcelAsMapByName(file, startRow, startCol, sheetName, headers, ReadStrategy.SAX);
    }
    
    /**
     * 按 sheet 名称读取 Excel 为 Map
     */
    public static List<Map<String, String>> readExcelAsMapByName(File file, int startRow, int startCol, 
                                                                  String sheetName, List<String> headers, 
                                                                  ReadStrategy strategy) {
        int sheetIndex = getSheetIndex(file, sheetName);
        if (sheetIndex < 0) {
            throw new ExcelReadException("找不到 Sheet: " + sheetName);
        }
        return readExcelAsMap(file, startRow, startCol, sheetIndex, headers, strategy);
    }
    
    // ==================== 读取图片 ====================
    
    /**
     * 读取指定 sheet 的所有图片（包含浮动图片和 DISPIMG 图片）
     * 
     * @param file Excel 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> readExcelImages(File file, int sheetNum) {
        return readExcelImages(file, sheetNum, ReadStrategy.SAX);
    }
    
    /**
     * 读取指定 sheet 的所有图片
     * 
     * @param file Excel 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @param strategy 读取策略
     * @return 图片列表
     */
    public static List<ImageData> readExcelImages(File file, int sheetNum, ReadStrategy strategy) {
        ExcelReader reader = getReader(strategy);
        return reader.readExcelImages(file, sheetNum);
    }
    
    /**
     * 按 sheet 名称读取所有图片
     */
    public static List<ImageData> readExcelImagesByName(File file, String sheetName) {
        return ImageParser.parse(file, sheetName);
    }
    
    /**
     * 按 sheet 名称和策略读取所有图片
     */
    public static List<ImageData> readExcelImagesByName(File file, String sheetName, ReadStrategy strategy) {
        return readExcelImagesByName(file, sheetName);
    }
    
    // ==================== 读取策略枚举 ====================
    
    public enum ReadStrategy {
        SAX,  // 流式读取，内存占用低，适合大文件
        DOM   // 文档对象模型，内存占用高，适合小文件
    }
}
