package com.lizhiwei.quickExcel.v3.reader;

import com.lizhiwei.quickExcel.core.ExcelReader;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.model.UploadFile;
import com.lizhiwei.quickExcel.v3.read.ReadExcelByXml;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * V3 引擎 Excel 读取器适配器
 * 包装 ReadExcelByXml 以适配统一的 ExcelReader 接口
 */
public class XmlExcelReaderAdapter implements ExcelReader {
    
    private final ReadExcelByXml.ReadStrategy strategy;
    
    public XmlExcelReaderAdapter() {
        this(ReadExcelByXml.ReadStrategy.SAX);
    }
    
    public XmlExcelReaderAdapter(ReadExcelByXml.ReadStrategy strategy) {
        this.strategy = strategy;
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, false, strategy);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe) {
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe, strategy);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe, boolean readImage) {
        // V3 引擎的图片读取是独立的，这里先忽略 readImage 参数
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe, strategy);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe, boolean readImage, 
                                  List<ExcelEntity> propertieList) {
        // V3 引擎暂时不支持自定义属性列表
        return ReadExcelByXml.readExcel(file, startRow, startCol, sheetNum, entity, safe, strategy);
    }
    
    @Override
    public List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, 
                                                     int sheetNum, boolean safe, 
                                                     List<ExcelEntity> propertieList) {
        return ReadExcelByXml.readExcelAsMap(file, startRow, startCol, sheetNum, null);
    }
    
    @Override
    public <T> List<T> readExcel(UploadFile file, int startRow, int startCol, 
                                  int sheetNum, Class<T> entity) {
        return readExcel(file.getFile(), startRow, startCol, sheetNum, entity);
    }
    
    @Override
    public <T> List<T> readExcel(UploadFile file, int startRow, int startCol, 
                                  int sheetNum, Class<T> entity, boolean safe) {
        return readExcel(file.getFile(), startRow, startCol, sheetNum, entity, safe);
    }
    
    @Override
    public <T> List<T> readExcel(UploadFile file, int startRow, int startCol, 
                                  int sheetNum, Class<T> entity, boolean safe, boolean readImage) {
        return readExcel(file.getFile(), startRow, startCol, sheetNum, entity, safe, readImage);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, String sheetName, 
                                  Class<T> entity, boolean safe, boolean readImage) {
        return ReadExcelByXml.readExcelByName(file, startRow, startCol, sheetName, entity, safe, strategy);
    }
}
