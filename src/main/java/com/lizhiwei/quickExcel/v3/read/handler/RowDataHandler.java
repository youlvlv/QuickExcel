package com.lizhiwei.quickExcel.v3.read.handler;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.v3.read.converter.ValueConverter;
import com.lizhiwei.quickExcel.v3.read.model.CellData;
import com.lizhiwei.quickExcel.v3.read.model.RowData;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 行数据处理器
 * 处理行数据的读取和转换
 */
public class RowDataHandler {
    
    /**
     * 将行数据转换为实体对象
     * @param rowData 行数据
     * @param columnMapping 列索引到属性的映射
     * @param entityClass 实体类
     * @param <T> 实体类型
     * @return 实体对象
     */
    public static <T> T toEntity(RowData rowData, Map<Integer, ExcelEntity> columnMapping, 
                                  Class<T> entityClass) throws Exception {
        T entity = ValueConverter.newInstance(entityClass);
        
        for (CellData cell : rowData.getCells()) {
            ExcelEntity property = columnMapping.get(cell.getColumnIndex());
            if (property != null) {
                String value = cell.getValue();
                ValueConverter.convertAndSet(entity, property, value);
            }
        }
        
        return entity;
    }
    
    /**
     * 将行数据转换为 Map
     * @param rowData 行数据
     * @param headers 表头列表
     * @param startCol 起始列索引
     * @param filterHeaders 需要过滤的表头列表（null 表示不过滤）
     * @return Map
     */
    public static Map<String, String> toMap(RowData rowData, List<String> headers,
                                           int startCol, List<String> filterHeaders) {
        Map<String, String> map = new LinkedHashMap<>();
        
        for (CellData cell : rowData.getCells()) {
            int colIndex = cell.getColumnIndex();
            if (colIndex < startCol) {
                continue;
            }
            
            String header = getHeader(headers, colIndex, startCol);
            
            if (filterHeaders == null || filterHeaders.isEmpty() || filterHeaders.contains(header)) {
                map.put(header, cell.getValue());
            }
        }
        
        return map;
    }
    
    /**
     * 获取表头
     */
    private static String getHeader(List<String> headers, int colIndex, int startCol) {
        int headerIndex = colIndex - startCol;
        if (headers != null && headerIndex < headers.size()) {
            String header = headers.get(headerIndex);
            if (header != null && !header.isEmpty()) {
                return header;
            }
        }
        return "Column" + (colIndex + 1);
    }
    
    /**
     * 解析行元素为 RowData
     * @param cells 单元格列表
     * @return RowData
     */
    public static RowData parse(List<CellData> cells) {
        RowData rowData = new RowData();
        
        if (cells != null && !cells.isEmpty()) {
            int rowIndex = cells.get(0).getRowIndex();
            rowData.setRowIndex(rowIndex);
            
            for (CellData cell : cells) {
                rowData.addCell(cell);
            }
        }
        
        return rowData;
    }
    
    /**
     * 获取表头行
     * @param allRows 所有行数据
     * @param headerRowIndex 表头行索引
     * @return 表头行的单元格列表
     */
    public static List<CellData> getHeaderRow(List<RowData> allRows, int headerRowIndex) {
        for (RowData row : allRows) {
            if (row.getRowIndex() == headerRowIndex) {
                return row.getCells();
            }
        }
        return new ArrayList<>();
    }
}
