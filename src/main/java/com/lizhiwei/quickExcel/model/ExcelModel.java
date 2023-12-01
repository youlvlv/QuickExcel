package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

/**
 * excel模型
 */
public class ExcelModel extends ExcelBaseModel implements AutoCloseable {

    /**
     * 默认的单元格格式
     */
    protected CellStyle DEFAULT_CELL_STYLE;

    protected XSSFWorkbook xWorkbook;

    protected CellStyle CustomCellStyle;

    /**
     * 获取poi模型
     *
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
        List<ExcelEntity> list = util().getExcelEntities(entity);
        util().setSheetHeader(sheet, list);
        util().setSheetContent(sheet, listContent, list);
        return this;
    }

    /**
     * 创建新 sheet
     *
     * @param name sheet 名
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
     *
     * @return
     */
    public CellStyle getDefaultStyle() {
        return DEFAULT_CELL_STYLE;
    }

    public ExcelModel setDefaultStyle(){
        return  this;
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
     *
     * @param operation 文件操作类
     * @return
     */
    public ExcelModel exportExcel(FileOperation operation) {
        operation.run(this);
        return this;
    }

    /**
     * 导出excel并关闭excel
     *
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
        DEFAULT_CELL_STYLE = xWorkbook.createCellStyle();
        //设置水平、垂直居中
        DEFAULT_CELL_STYLE.setAlignment(HorizontalAlignment.CENTER);
        DEFAULT_CELL_STYLE.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        Font headerFont = new XSSFFont();
        headerFont.setFontHeightInPoints((short) 12);
        /*headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);*/
        headerFont.setBold(true);
        headerFont.setFontName("宋体");
        DEFAULT_CELL_STYLE.setFont(headerFont);
        DEFAULT_CELL_STYLE.setWrapText(true);//是否自动换行
    }


    /**
     *
     * @param imagePath 图片路径
     * @param sheetModel 工作表
     * @param col1 // 图片起始列
     * @param row1 // 图片起始行
     * @param col2 // 图片结束列
     * @param row2 // 图片结束行
     * @throws IOException 异常
     */
    public void addPicture(String imagePath,SheetModel sheetModel,int col1,int row1,int col2,int row2) throws IOException {
        InputStream inputStream = new FileInputStream(imagePath);
        byte[] imageBytes = IOUtils.toByteArray(inputStream);
        inputStream.close();
        int pictureIdx=xWorkbook.addPicture(imageBytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
        CreationHelper helper = xWorkbook.getCreationHelper();
        Drawing drawing = sheetModel.createDrawingPatriarch();
        //添加位置
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(col1); // 图片起始列
        anchor.setRow1(row1); // 图片起始行
        anchor.setCol2(col2); // 图片结束列
        anchor.setRow2(row2); // 图片结束行
        // 创建图片并设置锚点
        Picture picture = drawing.createPicture(anchor, pictureIdx);
//        picture.resize();//自适应大小
    }
}

