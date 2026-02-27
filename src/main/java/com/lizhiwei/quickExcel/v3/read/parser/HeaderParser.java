package com.lizhiwei.quickExcel.v3.read.parser;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.v3.read.model.CellData;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 表头解析器
 * 用于解析 Excel 表头行，建立列索引与实体属性的映射关系
 */
public class HeaderParser {
    
    /**
     * 解析表头行，建立列索引到实体属性的映射
     * @param cells 表头行的单元格列表
     * @param properties 实体类的属性列表
     * @param startCol 起始列索引
     * @return 列索引到属性的映射
     */
    public static Map<Integer, ExcelEntity> parse(List<CellData> cells, 
                                                   List<ExcelEntity> properties, 
                                                   int startCol) {
        Map<Integer, ExcelEntity> mapping = new HashMap<>();
        
        for (CellData cell : cells) {
            int colIndex = cell.getColumnIndex();
            if (colIndex < startCol) {
                continue;
            }
            
            String headerValue = cell.getValue();
            if (headerValue == null || headerValue.isEmpty()) {
                continue;
            }
            
            for (ExcelEntity prop : properties) {
                if (isMatch(prop, headerValue)) {
                    mapping.put(colIndex, prop);
                    break;
                }
            }
        }
        
        return mapping;
    }
    
    /**
     * 解析表头行为字符串列表
     * @param cells 表头行的单元格列表
     * @param startCol 起始列索引
     * @return 表头字符串列表
     */
    public static List<String> parseAsList(List<CellData> cells, int startCol) {
        List<String> headers = new ArrayList<>();
        
        for (CellData cell : cells) {
            int colIndex = cell.getColumnIndex();
            if (colIndex < startCol) {
                continue;
            }
            
            while (headers.size() <= colIndex - startCol) {
                headers.add("");
            }
            headers.set(colIndex - startCol, cell.getValue());
        }
        
        return headers;
    }
    
    /**
     * 检查属性是否匹配表头值
     */
    private static boolean isMatch(ExcelEntity property, String headerValue) {
        if (property.getTitle() != null && property.getTitle().equals(headerValue)) {
            return true;
        }
        if (property.getAlias() != null && !property.getAlias().isEmpty() 
            && property.getAlias().equals(headerValue)) {
            return true;
        }
        return false;
    }
    
    /**
     * 如果没有表头行，按顺序建立默认映射
     * @param properties 实体类的属性列表
     * @param startCol 起始列索引
     * @return 列索引到属性的映射
     */
    public static Map<Integer, ExcelEntity> createDefaultMapping(List<ExcelEntity> properties, int startCol) {
        Map<Integer, ExcelEntity> mapping = new HashMap<>();
        int colIndex = startCol;
        
        for (ExcelEntity prop : properties) {
            mapping.put(colIndex++, prop);
        }
        
        return mapping;
    }
}
