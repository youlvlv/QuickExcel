package com.lizhiwei.quickExcel.v3.read.parser;

import com.lizhiwei.quickExcel.v3.read.model.CellData;

/**
 * 单元格引用解析器
 * 将单元格引用（如 "A1", "BC123"）解析为行索引和列索引
 */
public class CellRefParser {
    
    /**
     * 解析单元格引用
     * @param ref 单元格引用，如 "A1", "BC123"
     * @return 包含行索引和列索引的数组 [rowIndex, colIndex]
     */
    public static int[] parse(String ref) {
        if (ref == null || ref.isEmpty()) {
            return new int[] {0, 0};
        }
        
        StringBuilder colStr = new StringBuilder();
        StringBuilder rowStr = new StringBuilder();
        
        for (char c : ref.toCharArray()) {
            if (Character.isLetter(c)) {
                colStr.append(c);
            } else {
                rowStr.append(c);
            }
        }
        
        int col = columnToIndex(colStr.toString());
        int row = rowStr.length() > 0 ? Integer.parseInt(rowStr.toString()) : 0;
        
        return new int[] {row, col};
    }
    
    /**
     * 解析单元格引用到 CellData
     */
    public static CellData parseToCellData(String ref) {
        int[] coords = parse(ref);
        CellData cell = new CellData();
        cell.setRowIndex(coords[0]);
        cell.setColumnIndex(coords[1]);
        cell.setReference(ref);
        return cell;
    }
    
    /**
     * 将列字母转换为索引（0 基）
     * A -> 0, B -> 1, ..., Z -> 25, AA -> 26, AB -> 27, ...
     */
    public static int columnToIndex(String column) {
        if (column == null || column.isEmpty()) {
            return 0;
        }
        
        int col = 0;
        for (char c : column.toUpperCase().toCharArray()) {
            col = col * 26 + (c - 'A' + 1);
        }
        return col - 1;
    }
    
    /**
     * 将列索引转换为字母
     * 0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, 27 -> AB, ...
     */
    public static String indexToColumn(int index) {
        StringBuilder sb = new StringBuilder();
        int col = index + 1;
        
        while (col > 0) {
            col--;
            sb.insert(0, (char) ('A' + (col % 26)));
            col = col / 26;
        }
        
        return sb.toString();
    }
    
    /**
     * 将行索引和列索引转换为单元格引用
     */
    public static String toReference(int rowIndex, int colIndex) {
        return indexToColumn(colIndex) + rowIndex;
    }
}
