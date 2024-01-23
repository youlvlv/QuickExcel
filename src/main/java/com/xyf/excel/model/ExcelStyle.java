package com.xyf.excel.model;

import org.apache.poi.ss.usermodel.*;

public class ExcelStyle {

    /**
     * 表头设置
     * @param sheetModel
     * @return
     */
    public static CellStyle styleTitle(SheetModel sheetModel) {
        //创建表格的样式
        CellStyle cs = sheetModel.getExcel().getWorkbook().createCellStyle();
        //设置水平、垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        Font headerFont = sheetModel.getExcel().getWorkbook().createFont();
//        headerFont.setFontHeightInPoints((short) 11);
        headerFont.setBold(true);
        headerFont.setFontName("宋体");
        headerFont.setFontHeightInPoints((short) 20);
        cs.setFont(headerFont);
//        cs.setBorderBottom(BorderStyle.valueOf((short) 1));//边框
//        cs.setBorderLeft(BorderStyle.valueOf((short) 1));
//        cs.setBorderRight(BorderStyle.valueOf((short) 1));
//        cs.setBorderTop(BorderStyle.valueOf((short) 1));
        cs.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cs.setWrapText(true);//是否自动换行
        return cs;
    }
}
