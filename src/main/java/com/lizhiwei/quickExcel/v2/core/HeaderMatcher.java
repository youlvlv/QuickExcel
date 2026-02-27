package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.lizhiwei.quickExcel.v2.core.CellValueExtractor.getSimpleCellValue;
import static com.lizhiwei.quickExcel.v2.core.MergeCellResolver.getMergedRegionValue;

/**
 * 表头匹配器
 * 负责匹配 Excel 表头与实体类属性
 */
public class HeaderMatcher {
    
    /**
     * 匹配表头生成 ExcelEntity 列表
     * @param startRow 数据开始行（表头在 startRow - 1）
     * @param startCol 开始列
     * @param propertyList 实体类属性列表
     * @param sheet Sheet
     * @return 匹配后的属性列表
     */
    public static List<ExcelEntity> matchHeaders(int startRow, int startCol, 
                                                  List<ExcelEntity> propertyList, 
                                                  Sheet sheet) {
        List<ExcelEntity> matchedProperties = new ArrayList<>();
        
        // 获取表头行
        Row headerRow = sheet.getRow(startRow - 1);
        if (headerRow == null) {
            return matchedProperties;
        }
        
        int cellNum = headerRow.getLastCellNum();
        
        // 构建表头名称与位置映射
        Map<String, Integer> cellNameMap = new HashMap<>();
        for (int j = startCol; j < cellNum; j++) {
            String headerValue = getSimpleCellValue(getMergedRegionValue(sheet, startRow - 1, j));
            if (headerValue != null && !headerValue.isEmpty()) {
                cellNameMap.put(headerValue, j);
            }
        }
        
        // 匹配属性
        for (ExcelEntity excelEntity : propertyList) {
            if (excelEntity.isRead()) {
                Integer columnIndex = findColumnIndex(cellNameMap, excelEntity);
                if (columnIndex != null) {
                    excelEntity.setValue(columnIndex);
                    matchedProperties.add(excelEntity);
                }
            }
        }
        
        return matchedProperties;
    }
    
    /**
     * 查找列索引
     * 优先匹配 title，其次匹配 alias
     */
    private static Integer findColumnIndex(Map<String, Integer> cellNameMap, ExcelEntity excelEntity) {
        // 优先匹配 title
        if (cellNameMap.containsKey(excelEntity.getTitle())) {
            return cellNameMap.get(excelEntity.getTitle());
        }
        // 其次匹配 alias
        if (excelEntity.getAlias() != null && !excelEntity.getAlias().isEmpty() 
                && cellNameMap.containsKey(excelEntity.getAlias())) {
            return cellNameMap.get(excelEntity.getAlias());
        }
        return null;
    }
}
