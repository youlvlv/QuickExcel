package com.lizhiwei.quickExcel.model;


import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.Since;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetModel extends ExcelUtil {
    private final XSSFSheet xSheet;
    private final ExcelModel excel;

    protected int rowNum = 0;

    public SheetModel(XSSFSheet xSheet, ExcelModel excel) {
        this.xSheet = xSheet;
        this.excel = excel;
    }

    /**
     * 创建 数据信息
     *
     * @param entity
     * @param listContent
     * @param <T>
     * @return
     */
    public <T> SheetModel createInfo(Class<T> entity, List<T> listContent) {
        List<ExcelEntity> list = util.getExcelEntities(entity);
        SheetModel newSheet = util.setSheetHeader(this, list);
        return util.setSheetContent(newSheet, listContent, list, null);
    }

    public <T> SheetModel createHeader(Class<T> entity) {
        List<ExcelEntity> list = util.getExcelEntities(entity);
        return util.setSheetHeader(this, list);
    }

    public <T> SheetModel createContent(Class<T> entity, T content) {
        List<ExcelEntity> list = util.getExcelEntities(entity);
        List<T> first = new ArrayList<>();
        first.add(content);
        return util.setSheetContent(this, first, list, null);
    }

    public <T> SheetModel createContent(Class<T> entity, List<T> listContent) {
        List<ExcelEntity> list = util.getExcelEntities(entity);
        return util.setSheetContent(this, listContent, list, null);
    }

    public <T> SheetModel createContent(Class<T> entity, List<T> listContent, Since... since) {
        List<ExcelEntity> list = util.getExcelEntities(entity);
        return util.setSheetContent(this, listContent, list, Arrays.asList(since));
    }

    /**
     * 获取单行数据
     *
     * @return
     */
    public RowModel newRow() {
        return new RowModel(rowNum, xSheet.createRow(rowNum++), this);
    }

    /**
     * 新生成多行
     *
     * @return
     */
    public MoreRowModel newMoreRow() {
        return new MoreRowModel(rowNum, rowNum + 1, xSheet.createRow(rowNum++), xSheet.createRow(rowNum++), this);
    }

    public XSSFSheet getSheet() {
        return xSheet;
    }

    public ExcelModel getExcel() {
        return excel;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void addRowNum() {
        rowNum++;
    }


    /**
     * 结束本sheet编辑
     *
     * @return
     */
    public ExcelModel over() {
        return excel;
    }

    public void addMergedRegion(CellRangeAddress cellRangeAddress) {
        xSheet.addMergedRegion(cellRangeAddress);
    }
}
