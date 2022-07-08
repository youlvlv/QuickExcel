package com.lizhiwei.quickExcel.core;


import com.lizhiwei.quickExcel.entity.*;
import com.lizhiwei.quickExcel.exception.ExcelValueError;
import com.lizhiwei.quickExcel.model.MoreRowModel;
import com.lizhiwei.quickExcel.model.RowModel;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

public class ExcelUtil {

    protected static final ExcelUtil util = new ExcelUtil();

    /**
     * 根据 实体类生成 excel实体类
     *
     * @param entity
     * @param <T>
     * @return
     */
    public <T> List<ExcelEntity> getExcelEntities(Class<T> entity) {
        Field[] fields = entity.getDeclaredFields();
        List<ExcelEntity> listTitle = new ArrayList<>();
        int i = 0;
        for (Field field : fields) {
            //设置属性默认可访问，防止private阻止访问
            field.setAccessible(true);
            //判断是否包含Excel注解
            if (field.isAnnotationPresent(Excel.class)) {
                //获取Excel注解
                Excel e = field.getDeclaredAnnotation(Excel.class);
                int order = i++;
                ExcelEntity excelEntity;
                try {
                    excelEntity = new ExcelEntity(field.getName(), e.name().isEmpty() ? e.value() : e.name(), e.format().getDeclaredConstructor().newInstance(), e.index(), e.secondName());
                } catch (InvocationTargetException | InstantiationException | IllegalAccessException |
                         NoSuchMethodException ex) {
                    ex.printStackTrace();
                    excelEntity = new ExcelEntity(field.getName(), e.name().isEmpty() ? e.value() : e.name(), new DefaultFormat(), e.index(), DefaultTopName.class);
                }
                listTitle.add(excelEntity);
            }
        }
        listTitle.sort((o1, o2) -> {
            if (o2.getIndex() < 0) {
                return -1;
            }
            if (o1.getIndex() < 0) {
                return 1;
            }
            return Integer.compare(o1.getIndex(), o2.getIndex());
        });
        return listTitle;
    }

    /**
     * 生成数据头
     *
     * @param sheet
     * @param listTitle
     * @return
     */
    public SheetModel setSheetHeader(SheetModel sheet, List<ExcelEntity> listTitle) {
        //设置表格的宽度  xSheet.setColumnWidth(0, 20 * 256); 中的数字 20 自行设置为自己适用的
        /*xSheet.setColumnWidth(0, 20 * 256);
        xSheet.setColumnWidth(1, 15 * 256);
        xSheet.setColumnWidth(2, 15 * 256);
        xSheet.setColumnWidth(3, 20 * 256);*/

        //创建表格的样式
        CellStyle cs = sheet.getExcel().getWorkbook().createCellStyle();
        //设置水平、垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置字体
        Font headerFont = sheet.getExcel().getWorkbook().createFont();
        headerFont.setFontHeightInPoints((short) 12);
        /*headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);*/
        headerFont.setBold(true);
        headerFont.setFontName("宋体");
        cs.setFont(headerFont);
        cs.setWrapText(true);//是否自动换行
        //判断是否有多行头
        boolean moreRow = listTitle.stream().filter(x -> !x.getTopName().equals(DefaultTopName.class)).findAny().orElse(null) != null;
        if (moreRow) {
            MoreRowModel xRow0 = sheet.newMoreRow();
            //获取所有非默认头的字段
            Map<Class<? extends TopName>, List<ExcelEntity>> group = listTitle.stream().filter(x -> !x.getTopName().equals(DefaultTopName.class))
                    .collect(Collectors.groupingBy(ExcelEntity::getTopName));
//            Map<Integer,TopName> type = new HashMap<>();
            AtomicInteger order = new AtomicInteger();
            group.forEach((k, v) -> {
                try {
                    xRow0.setHeaderValue(order.get(), order.get() + v.size() - 1, k.getDeclaredConstructor().newInstance().value(), cs);
                    xRow0.setSecondHeaderValue(v);
                    order.getAndIncrement();
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                         NoSuchMethodException e) {
                    throw new RuntimeException(e);
                }
            });
            for (ExcelEntity excelEntity : listTitle) {
                if (excelEntity.getTopName().equals(DefaultTopName.class)) {
                    xRow0.setValue(order.getAndIncrement(), excelEntity.getTitle(), cs);
                }
            }
        } else {
            //创建一行
            RowModel xRow0 = sheet.newRow();
            int i = 0;
            for (ExcelEntity excelEntity : listTitle) {
                xRow0.setValue(i, excelEntity.getTitle(), cs);
//            XSSFCell xCell0 = xRow0.createCell(i);
//            xCell0.setCellStyle(cs);
//            xCell0.setCellValue(excelEntity.getTitle());
                i++;
            }
        }
        return sheet;
    }

    public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle) {
        return this.setSheetContent(sheet, listContent, listTitle, null);
    }


    /**
     * 配置(赋值)表格内容部分
     *
     * @param listContent
     * @param since
     * @throws Exception
     */
    public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since) {

        //创建内容样式（头部以下的样式）
        CellStyle cs = sheet.getExcel().getWorkbook().createCellStyle();
        cs.setWrapText(true);

        //设置水平垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        int start = sheet.getRowNum();
        if (null != listContent && listContent.size() > 0) {
            try {
                for (T t : listContent) {
                    RowModel xRow = sheet.newRow();
                    //获取类属性
                    Field field;
                    int order = 0;
                    for (ExcelEntity excelEntity : listTitle) {
                        String str = excelEntity.getProperty();
                        //获取完成get方法  首字母大写如：getId
                        field = t.getClass().getDeclaredField(str);
                        field.setAccessible(true);
                        Object o = field.get(t);
                        String value = "";
                        ExcelFormat format = excelEntity.getFormat();
                        value = format.WriterToExcel(o);
                        //循环设置每列的值
                        xRow.setValue(order++, value, cs);
//                        XSSFCell xCell = xRow.createCell(j);
//                        xCell.setCellStyle(cs);
//                        xCell.setCellValue(value);
                    }
                }
                if (since != null) {
                    for (Since s : since) {
                        int i = listTitle.stream().filter(x -> x.getProperty().equals(s.getTitle())).findFirst().get().getIndex();
                        sheet.addMergedRegion(new CellRangeAddress(start, sheet.getRowNum() - 1, i, i));
                    }
                }
            } catch (IllegalAccessException | NoSuchFieldException e) {
                throw new ExcelValueError(e);
            }
        }
        return sheet;
    }

}
