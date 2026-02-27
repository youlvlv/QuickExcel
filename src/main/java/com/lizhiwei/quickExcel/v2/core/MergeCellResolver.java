package com.lizhiwei.quickExcel.v2.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 合并单元格解析器
 * 负责处理 Excel 中的合并单元格
 */
public class MergeCellResolver {
    
    /**
     * 获取合并单元格的值
     * 如果指定位置在合并区域内，返回合并区域左上角单元格的值
     * @param sheet Sheet
     * @param row 行号
     * @param column 列号
     * @return 单元格对象
     */
    public static Cell getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    return fRow.getCell(firstColumn);
                }
            }
        }
        
        Row targetRow = sheet.getRow(row);
        return targetRow != null ? targetRow.getCell(column) : null;
    }
    
    /**
     * 判断指定位置是否在合并区域内
     */
    public static boolean isInMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            if (row >= ca.getFirstRow() && row <= ca.getLastRow()
                    && column >= ca.getFirstColumn() && column <= ca.getLastColumn()) {
                return true;
            }
        }
        return false;
    }
}
