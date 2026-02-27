package com.lizhiwei.quickExcel.v3.read.parser;

import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;

import java.io.File;
import java.util.List;

/**
 * 图片解析器
 * 用于从 xlsx 文件中解析图片
 * <p>
 * 现在基于 XML 直接解析，不再依赖 XSSFWorkbook
 * </p>
 */
public class ImageParser {
    
    /**
     * 从 xlsx 文件中解析图片（使用上下文）
     * @param context 文件上下文
     * @param sheetNum sheet 索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> parse(ExcelFileContext context, int sheetNum) {
        return XmlImageParser.parse(context, sheetNum);
    }
    
    /**
     * 从 xlsx 文件中解析图片（使用上下文，按 sheet 名称）
     * @param context 文件上下文
     * @param sheetName sheet 名称
     * @return 图片列表
     */
    public static List<ImageData> parse(ExcelFileContext context, String sheetName) {
        return XmlImageParser.parse(context, sheetName);
    }
    
    /**
     * 从 xlsx 文件中解析图片
     * @param file xlsx 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> parse(File file, int sheetNum) {
        // 使用 XmlImageParser 直接解析 XML
        return XmlImageParser.parse(file, sheetNum);
    }
    
    /**
     * 从 xlsx 文件中解析图片（按 sheet 名称）
     * @param file xlsx 文件
     * @param sheetName sheet 名称
     * @return 图片列表
     */
    public static List<ImageData> parse(File file, String sheetName) {
        return XmlImageParser.parse(file, sheetName);
    }
    
    /**
     * 根据行号和列号获取图片
     * @param file xlsx 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @param row 行索引（从 0 开始）
     * @param column 列索引（从 0 开始）
     * @return 图片数据
     */
    public static ImageData getImage(File file, int sheetNum, int row, int column) {
        return XmlImageParser.getImage(file, sheetNum, row, column);
    }
    
    /**
     * 根据行号和列号获取图片（按 sheet 名称）
     * @param file xlsx 文件
     * @param sheetName sheet 名称
     * @param row 行索引（从 0 开始）
     * @param column 列索引（从 0 开始）
     * @return 图片数据
     */
    public static ImageData getImage(File file, String sheetName, int row, int column) {
        return XmlImageParser.getImage(file, sheetName, row, column);
    }
    
    /**
     * 根据行号获取该行的所有图片
     * @param file xlsx 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @param row 行索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> getImagesByRow(File file, int sheetNum, int row) {
        return XmlImageParser.getImagesByRow(file, sheetNum, row);
    }
    
    /**
     * 根据行号获取该行的所有图片（按 sheet 名称）
     * @param file xlsx 文件
     * @param sheetName sheet 名称
     * @param row 行索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> getImagesByRow(File file, String sheetName, int row) {
        return XmlImageParser.getImagesByRow(file, sheetName, row);
    }
    
    /**
     * 根据 MIME 类型获取文件扩展名
     */
    public static String getExtension(String mimeType) {
        return XmlImageParser.getExtension(mimeType);
    }
}
