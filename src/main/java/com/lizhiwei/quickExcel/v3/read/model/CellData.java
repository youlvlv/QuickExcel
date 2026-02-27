package com.lizhiwei.quickExcel.v3.read.model;

/**
 * 单元格数据模型
 */
public class CellData {
    
    private int rowIndex;
    private int columnIndex;
    private String reference;
    private String value;
    private String rawValue;
    private CellType cellType;
    private Integer styleIndex;
    private String formatString;
    
    public enum CellType {
        STRING,
        INLINE_STRING,
        SHARED_STRING,
        NUMERIC,
        BOOLEAN,
        ERROR,
        FORMULA,
        EMPTY
    }
    
    public CellData() {
    }
    
    public CellData(int rowIndex, int columnIndex, String value) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.value = value;
    }
    
    public int getRowIndex() {
        return rowIndex;
    }
    
    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
    
    public int getColumnIndex() {
        return columnIndex;
    }
    
    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }
    
    public String getReference() {
        return reference;
    }
    
    public void setReference(String reference) {
        this.reference = reference;
    }
    
    public String getValue() {
        return value;
    }
    
    public void setValue(String value) {
        this.value = value;
    }
    
    public String getRawValue() {
        return rawValue;
    }
    
    public void setRawValue(String rawValue) {
        this.rawValue = rawValue;
    }
    
    public CellType getCellType() {
        return cellType;
    }
    
    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }
    
    public Integer getStyleIndex() {
        return styleIndex;
    }
    
    public void setStyleIndex(Integer styleIndex) {
        this.styleIndex = styleIndex;
    }
    
    public String getFormatString() {
        return formatString;
    }
    
    public void setFormatString(String formatString) {
        this.formatString = formatString;
    }
}
