package com.lizhiwei.quickExcel.v3.read.handler;

import com.lizhiwei.quickExcel.v3.read.model.CellData;
import com.lizhiwei.quickExcel.v3.read.parser.SharedStringParser;

import java.util.List;

/**
 * 单元格数据处理器
 * 处理单元格值的读取和类型转换
 */
public class CellDataHandler {
    
    /**
     * 处理单元格值
     * @param cell 单元格数据
     * @param sharedStrings 共享字符串列表
     * @return 处理后的值
     */
    public static String process(CellData cell, List<String> sharedStrings) {
        if (cell == null) {
            return "";
        }
        
        CellData.CellType cellType = cell.getCellType();
        
        switch (cellType) {
            case SHARED_STRING:
                return processSharedString(cell, sharedStrings);
            case INLINE_STRING:
                return processInlineString(cell);
            case NUMERIC:
            case BOOLEAN:
            case ERROR:
                return cell.getRawValue() != null ? cell.getRawValue() : "";
            case FORMULA:
                return cell.getRawValue() != null ? cell.getRawValue() : "";
            case STRING:
                return cell.getValue() != null ? cell.getValue() : "";
            case EMPTY:
            default:
                return "";
        }
    }
    
    /**
     * 处理共享字符串
     */
    private static String processSharedString(CellData cell, List<String> sharedStrings) {
        String rawValue = cell.getRawValue();
        if (rawValue == null || rawValue.isEmpty()) {
            return "";
        }
        
        try {
            int index = Integer.parseInt(rawValue);
            return SharedStringParser.getString(sharedStrings, index);
        } catch (NumberFormatException e) {
            return rawValue;
        }
    }
    
    /**
     * 处理内联字符串
     */
    private static String processInlineString(CellData cell) {
        return cell.getValue() != null ? cell.getValue() : "";
    }
    
    /**
     * 创建空单元格数据
     */
    public static CellData createEmpty(int rowIndex, int colIndex) {
        CellData cell = new CellData();
        cell.setRowIndex(rowIndex);
        cell.setColumnIndex(colIndex);
        cell.setCellType(CellData.CellType.EMPTY);
        cell.setValue("");
        return cell;
    }
}
