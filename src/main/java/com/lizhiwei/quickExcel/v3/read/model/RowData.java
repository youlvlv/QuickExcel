package com.lizhiwei.quickExcel.v3.read.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 行数据模型
 * 存储一行的所有单元格数据
 */
public class RowData {
    
    private int rowIndex;
    private final List<CellData> cells;
    
    public RowData(int rowIndex) {
        this.rowIndex = rowIndex;
        this.cells = new ArrayList<>();
    }
    
    public RowData() {
        this.cells = new ArrayList<>();
    }
    
    public void addCell(CellData cell) {
        cells.add(cell);
    }
    
    public CellData getCell(int columnIndex) {
        return cells.stream()
                .filter(c -> c.getColumnIndex() == columnIndex)
                .findFirst()
                .orElse(null);
    }
    
    public String getCellValue(int columnIndex) {
        CellData cell = getCell(columnIndex);
        return cell != null ? cell.getValue() : null;
    }
    
    public int getRowIndex() {
        return rowIndex;
    }
    
    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
    
    public List<CellData> getCells() {
        return cells;
    }
    
    public boolean isEmpty() {
        return cells.stream().allMatch(c -> c.getValue() == null || c.getValue().isEmpty());
    }
}
