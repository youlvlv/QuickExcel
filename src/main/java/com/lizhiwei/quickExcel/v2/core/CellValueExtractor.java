package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;

/**
 * 单元格值提取器
 * 负责从 POI Cell 中提取各种类型的值
 */
public class CellValueExtractor {
    
    private static final SimpleDateFormat DEFAULT_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");
    
    /**
     * 获取单元格的字符串值
     * @param workbook Workbook
     * @param cell 单元格
     * @param property Excel 实体属性
     * @return 字符串值
     */
    public static String getCellStringValue(Workbook workbook, Cell cell, ExcelEntity property) {
        String cellValue = "";
        if (cell != null) {
            cellValue = extractCellValue(workbook, cell, property.getAccuracy());
            // 判断当前字段是否允许非空
            if (property.isNotNull() && cellValue.trim().isEmpty()) {
                throw new ExcelValueException(property.getTitle() + "为空");
            }
            if (!cellValue.trim().isEmpty()) {
                return cellValue;
            }
        }
        return "";
    }
    
    /**
     * 获取单元格的字符串值（带格式化）
     * @param workbook Workbook
     * @param cell 单元格
     * @param property Excel 实体属性
     * @return 格式化后的字符串值
     */
    public static String getExcelStringValue(Workbook workbook, Cell cell, ExcelEntity property) {
        if (cell == null) {
            return null;
        }
        
        String cellValue = extractCellValue(workbook, cell, property.getAccuracy());
        
        // 判断当前字段是否允许非空
        if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
            throw new ExcelValueException(property.getTitle() + "为空");
        }
        
        if (!cellValue.trim().isEmpty()) {
            return applyFormat(property, cellValue);
        }
        return null;
    }
    
    /**
     * 提取单元格原始值
     */
    private static String extractCellValue(Workbook workbook, Cell cell, int accuracy) {
        return extractCellValue(workbook, cell, "", DEFAULT_DATE_FORMAT, accuracy);
    }
    
    /**
     * 提取单元格原始值（完整版）
     */
    private static String extractCellValue(Workbook workbook, Cell cell, String defaultValue, 
                                           SimpleDateFormat dateFormat, int accuracy) {
        if (cell == null) {
            return defaultValue;
        }
        
        return switch (cell.getCellType()) {
            case NUMERIC -> handleNumericCell(cell, dateFormat, accuracy);
            case STRING -> cell.getStringCellValue();
            case BLANK -> "";
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> handleFormulaCell(workbook, cell);
            case ERROR -> String.valueOf(cell.getErrorCellValue());
            default -> defaultValue;
        };
    }
    
    /**
     * 处理数值类型单元格
     */
    private static String handleNumericCell(Cell cell, SimpleDateFormat dateFormat, int accuracy) {
        if (DateUtil.isCellDateFormatted(cell)) {
            return dateFormat.format(cell.getDateCellValue());
        }
        
        double numericValue = cell.getNumericCellValue();
        
        // 判断是否为数学意义上的整数
        if (numericValue == Math.floor(numericValue) && !Double.isInfinite(numericValue)) {
            return String.valueOf((long) numericValue);
        }
        
        // 有小数部分
        if (accuracy == -1) {
            return new BigDecimal(String.valueOf(numericValue))
                    .stripTrailingZeros()
                    .toPlainString();
        } else {
            return BigDecimal.valueOf(numericValue)
                    .setScale(accuracy, RoundingMode.HALF_UP)
                    .toPlainString();
        }
    }
    
    /**
     * 处理公式类型单元格
     */
    private static String handleFormulaCell(Workbook workbook, Cell cell) {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        CellValue evaluated = evaluator.evaluate(cell);
        
        return switch (evaluated.getCellType()) {
            case NUMERIC -> String.valueOf(evaluated.getNumberValue());
            case STRING -> evaluated.getStringValue();
            default -> "";
        };
    }
    
    /**
     * 应用格式化
     */
    private static String applyFormat(ExcelEntity property, String cellValue) {
        if (property.getFormat() == null) {
            return cellValue;
        }
        
        ExcelFormatBase<?> format = property.getFormat();
        try {
            if (format instanceof DefaultFormat) {
                return ((DefaultFormat) format).ReadToExcel(String.class, cellValue).toString();
            }
            return format.ReadToExcel(cellValue, null).toString();
        } catch (Exception e) {
            throw new ExcelValueException(property.getTitle() + "错误", e);
        }
    }
    
    /**
     * 获取单元格的简单字符串表示
     */
    public static String getSimpleCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        return cell.toString();
    }
}
