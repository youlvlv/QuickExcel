package com.lizhiwei.quickExcel.v3.read.reader;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.util.ReadExcel;
import com.lizhiwei.quickExcel.v3.read.ReadExcelByXml;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;
import com.lizhiwei.quickExcel.v3.read.model.ExcelReadResult;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * Excel 读取器抽象基类
 */
public abstract class AbstractExcelReader implements ExcelReader {
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int sheetNum, Class<T> entity) {
        return readExcel(file, startRow, 0, sheetNum, entity, false);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return readExcel(file, startRow, startCol, sheetNum, entity, false);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe) {
        List<ExcelEntity> properties = getExcelEntities(entity);
        var list = doReadExcel(file, startRow, startCol, sheetNum, entity, properties, safe);
        ExcelFileContextManager.getInstance().removeContext(file);
        return list;
    }
    
    @Override
    public List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, int sheetNum, List<String> headers) {
        var list =  doReadExcelAsMap(file, startRow, startCol, sheetNum, headers);
        ExcelFileContextManager.getInstance().removeContext(file);
        return list;
    }
    
    /**
     * 一步读取 Excel 实体类和图片
     * 
     * @param file Excel 文件
     * @param startRow 起始行（从 1 开始）
     * @param startCol 起始列（从 0 开始）
     * @param sheetNum sheet 索引（从 0 开始）
     * @param entity 实体类
     * @param safe 是否安全模式
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    public <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int startCol, 
                                                       int sheetNum, Class<T> entity, boolean safe) {
        List<ExcelEntity> properties = getExcelEntities(entity);
        var result = doReadExcelWithImages(file, startRow, startCol, sheetNum, entity, properties, safe);
        ExcelFileContextManager.getInstance().removeContext(file);
        return result;

    }
    
    /**
     * 一步读取 Excel 实体类和图片（简化版）
     */
    public <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int sheetNum, Class<T> entity) {
        return readExcelWithImages(file, startRow, 0, sheetNum, entity, false);
    }
    
    /**
     * 一步读取 Excel 实体类和图片（带起始列）
     */
    public <T> ExcelReadResult<T> readExcelWithImages(File file, int startRow, int startCol, 
                                                       int sheetNum, Class<T> entity) {
        return readExcelWithImages(file, startRow, startCol, sheetNum, entity, false);
    }
    
    /**
     * 获取实体类的 ExcelEntity 列表
     */
    protected List<ExcelEntity> getExcelEntities(Class<?> entity) {
        return ReadExcel.getExcelEntities(entity);
    }
    
    /**
     * 验证 Sheet 索引
     */
    protected void validateSheetIndex(int sheetNum, int actualSheetCount) {
        if (sheetNum < 0 || sheetNum >= actualSheetCount) {
            throw new ExcelReadException("Sheet 索引超出范围: " + sheetNum);
        }
    }
    
    /**
     * 执行读取 Excel 的具体实现
     */
    protected abstract <T> List<T> doReadExcel(File file, int startRow, int startCol, int sheetNum, 
                                               Class<T> entity, List<ExcelEntity> properties, boolean safe);
    
    /**
     * 执行读取 Excel 为 Map 列表的具体实现
     */
    protected abstract List<Map<String, String>> doReadExcelAsMap(File file, int startRow, int startCol, 
                                                                   int sheetNum, List<String> headers);
    
    /**
     * 执行一步读取 Excel 实体类和图片的具体实现
     * 
     * @param file Excel 文件
     * @param startRow 起始行
     * @param startCol 起始列
     * @param sheetNum sheet 索引
     * @param entity 实体类
     * @param properties 实体属性列表
     * @param safe 是否安全模式
     * @param <T> 实体类型
     * @return 包含实体类和图片的读取结果
     */
    protected abstract <T> ExcelReadResult<T> doReadExcelWithImages(File file, int startRow, int startCol, 
                                                                     int sheetNum, Class<T> entity, 
                                                                     List<ExcelEntity> properties, boolean safe);
}
