package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * excel模型
 */
public class ExcelModel extends ExcelUtil implements AutoCloseable {
    protected XSSFWorkbook xWorkbook;

    /**
     * 获取poi模型
     * @return
     */
    public XSSFWorkbook getWorkbook() {
        return xWorkbook;
    }

    /**
     * 通过流写出
     *
     * @param stream
     */
    public void write(OutputStream stream) throws IOException {
        xWorkbook.write(stream);
    }

    /**
     * 创建新 sheet
     * 并填充数据
     *
     * @param name        sheet 名
     * @param entity      实体类
     * @param listContent 内容
     * @param <T>         实体类
     * @return
     */
    public <T> ExcelModel newSheet(String name, Class<T> entity, List<T> listContent) {
        SheetModel sheet = this.newSheet(name);
        List<ExcelEntity> list = util.getExcelEntities(entity);
        util.setSheetHeader(sheet, list);
        util.setSheetContent(sheet, listContent, list);
        return this;
    }

    /**
     * 创建新 sheet
     *
     * @param name        sheet 名
     * @return
     */
    public SheetModel newSheet(String name) {
        XSSFSheet xSheet = xWorkbook.createSheet(name);
        return new SheetModel(xSheet, this);
    }


    /**
     * 创建新 sheet
     *
     * @return
     */
    public SheetModel newSheet() {
        XSSFSheet xSheet = xWorkbook.createSheet();
        return new SheetModel(xSheet, this);
    }

    /**
     * 获取默认单元格格式
     * @return
     */
    public CellStyle getDefaultStyle() {
        CellStyle cs = xWorkbook.createCellStyle();
        //设置水平、垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        Font headerFont = xWorkbook.createFont();
        headerFont.setFontHeightInPoints((short) 12);
        /*headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);*/
        headerFont.setBold(true);
        headerFont.setFontName("宋体");
        cs.setFont(headerFont);
        cs.setWrapText(true);//是否自动换行
        return cs;
    }

//    public ExcelModel exportExcel(FileOutputStream stream) {
//        try {
//            xWorkbook.write(stream);
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }
//        return this;
//    }

    /**
     * 导出excel
     * @param operation 文件操作类
     * @return
     */
    public ExcelModel exportExcel(FileOperation operation) {
        operation.run(this);
        return this;
    }

    /**
     * 导出excel并关闭excel
     * @param operation
     */
    public void exportExcelAndClose(FileOperation operation) {
        operation.run(this);
        this.close();
    }

    public void close() {
        try {
            xWorkbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {

        }
    }

    public ExcelModel() {
        xWorkbook = new XSSFWorkbook();
    }
}

