package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.exception.IORunTimeException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Workbook 加载器
 * 负责加载 Excel 文件为 POI Workbook 对象
 */
public class WorkbookLoader {
    
    /**
     * 加载 Workbook
     * @param file Excel 文件
     * @return Workbook 对象
     * @throws IOException 读取失败时抛出
     */
    public static Workbook load(File file) throws IOException {
        try (FileInputStream fi = new FileInputStream(file)) {
            String fileType = getFileExtension(file);
            return switch (fileType) {
                case "xls" -> new HSSFWorkbook(fi);
                case "xlsx" -> new XSSFWorkbook(fi);
                default -> throw new IORunTimeException("您导入的文件不是标准excel文件");
            };
        }
    }
    
    /**
     * 获取文件扩展名
     */
    private static String getFileExtension(File file) {
        String fileName = file.getName();
        int lastDotIndex = fileName.lastIndexOf(".");
        return lastDotIndex > 0 ? fileName.substring(lastDotIndex + 1).toLowerCase() : "";
    }
}
