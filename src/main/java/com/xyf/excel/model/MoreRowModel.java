package com.xyf.excel.model;

import com.xyf.excel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.List;

public class MoreRowModel extends RowModel {

    /**
     * 结束行数
     */
    protected final int endRowNumber;


    protected final XSSFRow secondRow;


    public MoreRowModel(int rowNumber, int endRowNumber, XSSFRow row, XSSFRow secondRow, SheetModel sheetModel) {
        super(rowNumber, row, sheetModel);
        this.endRowNumber = endRowNumber;
        this.secondRow = secondRow;
    }


    public RowModel setValue(int i, String firstValue, String secondValue, CellStyle style) {
        XSSFCell cell = row.createCell(i);
        cell.setCellValue(firstValue);
        cell.setCellStyle(style);
        XSSFCell cell2 = secondRow.createCell(i);
        cell2.setCellValue(secondValue);
        cell2.setCellStyle(style);
        return this;
    }

    public RowModel setHeaderValue(int i, int end, String value, CellStyle cs) {
        sheet.addMergedRegion(new CellRangeAddress(rowNumber, rowNumber, i, end));
        XSSFCell cell = row.createCell(i);
        cell.setCellValue(value);
        cell.setCellStyle(cs);
        return this;
    }

    @Override
    public RowModel setValue(int i, String value, CellStyle style) {
        sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, i, i));
        XSSFCell cell = createCell(i);
        cell.setCellValue(value);
        cell.setCellStyle(style);
        return this;
    }

    @Override
    public RowModel setMergerValue(int start, int end, String value) {
        sheet.addMergedRegion(new CellRangeAddress(rowNumber, endRowNumber, start, end));
        XSSFCell cell = createCell(start);
        cell.setCellValue(value);
        cell.setCellStyle(sheet.getExcel().getDefaultStyle());
        return this;
    }


    public void setSecondHeaderValue(List<ExcelEntity> v,CellStyle style) {
        for (ExcelEntity excelEntity : v) {
            XSSFCell cell = secondRow.createCell(excelEntity.getIndex());
            cell.setCellValue(excelEntity.getTitle());
            cell.setCellStyle(style);
        }
    }
}
