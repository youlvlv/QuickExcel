package com.lizhiwei.quickExcel.v2.reader;

import com.lizhiwei.quickExcel.core.ExcelReader;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.PictureMap;
import com.lizhiwei.quickExcel.entity.ReadErrorInfo;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.model.ExcelBaseModel;
import com.lizhiwei.quickExcel.model.UploadFile;
import com.lizhiwei.quickExcel.v2.core.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

import static com.lizhiwei.quickExcel.v2.core.CellValueExtractor.getCellStringValue;
import static com.lizhiwei.quickExcel.v2.core.CellValueExtractor.getExcelStringValue;
import static com.lizhiwei.quickExcel.v2.core.EntityPopulator.populate;
import static com.lizhiwei.quickExcel.v2.core.EntityValueConverter.convert;
import static com.lizhiwei.quickExcel.v2.core.HeaderMatcher.matchHeaders;
import static com.lizhiwei.quickExcel.v2.core.MergeCellResolver.getMergedRegionValue;

/**
 * V2 引擎 Excel 读取器
 * 基于 Apache POI 实现
 */
public class PoiExcelReader implements ExcelReader {
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity) {
        return readExcel(file, startRow, startCol, sheetNum, entity, false, false);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe) {
        return readExcel(file, startRow, startCol, sheetNum, entity, safe, false);
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe, boolean readImage) {
        return readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage, 
                         ExcelBaseModel.getExcelEntities(entity));
    }
    
    @Override
    public <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, 
                                  Class<T> entity, boolean safe, boolean readImage, 
                                  List<ExcelEntity> propertieList) {
        List<T> result = new ArrayList<>();
        boolean hasError = false;
        List<ReadErrorInfo> errorInfoList = new ArrayList<>();
        
        try (Workbook wb = WorkbookLoader.load(file)) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            PictureMap pictureMap = readImage ? ImageExtractor.extractPictures(sheet) : null;
            
            List<ExcelEntity> properties = matchHeaders(startRow, startCol, propertieList, sheet);
            if (properties.isEmpty()) {
                throw new ExcelReadException("未匹配到任何列，请检查表头是否正确");
            }
            
            int rowNum = sheet.getLastRowNum() + 1;
            EmptyRowChecker emptyChecker = new EmptyRowChecker();
            
            for (int i = startRow; i < rowNum; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                
                T instance = createInstance(entity);
                int nonEmptyCount = properties.size();
                Map<String, String> objectMap = new HashMap<>();
                Map<String, ExcelEntity> propertyMap = new HashMap<>();
                
                // 读取单元格值
                for (ExcelEntity property : properties) {
                    try {
                        String value = readCellValue(wb, sheet, i, property);
                        objectMap.put(property.getProperty(), value);
                        
                        // 处理别名
                        if (!Objects.equals(property.getAliasProperty(), "") 
                                && !property.getAliasProperty().equals(property.getProperty())) {
                            if (objectMap.containsKey(property.getAliasProperty())) {
                                throw new ExcelReadException("存在重复的 property: " + property.getAliasProperty());
                            }
                            objectMap.put(property.getAliasProperty(), value);
                        }
                        
                        propertyMap.put(property.getProperty(), property);
                        if (value.isEmpty()) {
                            nonEmptyCount--;
                        }
                    } catch (ExcelValueException e) {
                        if (safe) {
                            hasError = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + i + "行 " + e.getMessage(), e);
                        }
                    }
                }
                
                // 设置实体值
                for (Map.Entry<String, ExcelEntity> entry : propertyMap.entrySet()) {
                    try {
                        ExcelEntity property = entry.getValue();
                        Object value = convert(objectMap.get(entry.getKey()), property, objectMap);
                        populate(instance, property, value, pictureMap, row.getRowNum());
                    } catch (ExcelValueException e) {
                        if (safe) {
                            hasError = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + (i + startRow) + "行 " + e.getMessage(), e);
                        }
                    }
                }
                
                // 检查空行
                if (emptyChecker.checkEmptyRow(nonEmptyCount)) {
                    break;
                }
                if (nonEmptyCount > 0) {
                    result.add(instance);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException("读取 Excel 文件失败", e);
        }
        
        if (hasError) {
            throw new ExcelReadException(errorInfoList);
        }
        if (result.isEmpty()) {
            throw new ExcelReadException("当前表格为空");
        }
        
        return result;
    }
    
    @Override
    public List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, 
                                                     int sheetNum, boolean safe, 
                                                     List<ExcelEntity> propertieList) {
        List<Map<String, String>> result = new ArrayList<>();
        boolean hasError = false;
        List<ReadErrorInfo> errorInfoList = new ArrayList<>();
        
        try (Workbook wb = WorkbookLoader.load(file)) {
            Sheet sheet = wb.getSheetAt(sheetNum);
            List<ExcelEntity> properties = matchHeaders(startRow, startCol, propertieList, sheet);
            
            int rowNum = sheet.getLastRowNum() + 1;
            EmptyRowChecker emptyChecker = new EmptyRowChecker();
            
            for (int i = startRow; i < rowNum; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                
                Map<String, String> rowMap = new HashMap<>();
                int nonEmptyCount = properties.size();
                
                for (ExcelEntity property : properties) {
                    try {
                        String value = getExcelStringValue(wb, 
                            getMergedRegionValue(sheet, i, property.getValue()), property);
                        if (value == null || value.isEmpty()) {
                            nonEmptyCount--;
                        }
                        rowMap.put(property.getProperty(), value);
                    } catch (ExcelValueException e) {
                        if (safe) {
                            hasError = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + i + "行 " + e.getMessage(), e);
                        }
                    }
                }
                
                if (emptyChecker.checkEmptyRow(nonEmptyCount)) {
                    break;
                }
                if (nonEmptyCount > 0) {
                    result.add(rowMap);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException("读取 Excel 文件失败", e);
        }
        
        if (hasError) {
            throw new ExcelReadException(errorInfoList);
        }
        if (result.isEmpty()) {
            throw new ExcelReadException("当前表格为空");
        }
        
        return result;
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
        try (Workbook wb = WorkbookLoader.load(file)) {
            int sheetNum = wb.getSheetIndex(wb.getSheet(sheetName));
            return readExcel(file, startRow, startCol, sheetNum, entity, safe, readImage);
        } catch (IOException e) {
            throw new RuntimeException("读取 Excel 文件失败", e);
        }
    }
    
    /**
     * 读取单元格值
     */
    private String readCellValue(Workbook wb, Sheet sheet, int rowNum, ExcelEntity property) {
        return getCellStringValue(wb, getMergedRegionValue(sheet, rowNum, property.getValue()), property);
    }
    
    /**
     * 创建实体实例
     */
    private <T> T createInstance(Class<T> entity) {
        try {
            return entity.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException | 
                 NoSuchMethodException e) {
            throw new RuntimeException("构建实体类失败！请检查实体类是否有无参构造器", e);
        }
    }
}
