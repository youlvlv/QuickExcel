package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.entity.DefaultFormat;
import com.lizhiwei.quickExcel.entity.Excel;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ExcelFormat;
import com.lizhiwei.quickExcel.exception.ExcelValueError;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ReadExcel {


    public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity) {
        List<T> varList = new ArrayList<>();

        try {
            //读取文件
            FileInputStream fi = new FileInputStream(file);
            String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
            Workbook wb = null;
            //判断文件类型
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook(fi);
            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(fi);
            } else {
                throw new IORunTimeException("您导入的文件不是标准excel文件");
            }
            Sheet sheet = wb.getSheetAt(sheetnum); // sheet 从0开始
            List<ExcelEntity> properties = new ArrayList<>();
            //获取实体类所有属性
            Field[] fields = entity.getDeclaredFields();
            /*----------匹配头------------*/
            Row row = sheet.getRow(startrow - 1); // 行
            int cellNum = row.getLastCellNum(); // 每行的最后一个单元格位置
            //首行名称与位置
            Map<String, Integer> cellName = new HashMap<>();
            for (int j = startcol; j < cellNum; j++) { // 列循环开始
                cellName.put(getCellValue(getMergedRegionValue(sheet, startrow - 1, j)), j);
            }
            //循环实体类所有属性
            for (Field field : fields) {
                field.setAccessible(true);
                //判断当前字段是否被excel注解
                if (!field.isAnnotationPresent(Excel.class)) {
                    continue;
                }
                //生成实体类
                ExcelEntity excelEntity = new ExcelEntity();
                Excel excel = field.getAnnotation(Excel.class);
                //匹配是否为相同头
                if (cellName.containsKey(excel.value())) {
                    excelEntity.setProperty(field.getName());
                    excelEntity.setValue(cellName.get(excel.value()));
                    //实体类中该属性类型
                    excelEntity.setType(field.getType());
                    try {
                        excelEntity.setFormat(excel.format().getDeclaredConstructor().newInstance());
                    } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                             NoSuchMethodException e) {
                        e.printStackTrace();
                        //当前转换器生成实例时抛出异常，使用默认转换器
                        excelEntity.setFormat(new DefaultFormat());
                    }
                    properties.add(excelEntity);
                }
            }
            int rowNum = sheet.getLastRowNum() + 1; // 取得最后一行的行号
            //空行数
            int emptySize = 0;
            /*--------------数据行-----------------------*/
            for (int i = startrow; i < rowNum; i++) { // 行循环开始

                row = sheet.getRow(i); // 行
                if (row == null) {
                    break;
                }
                T t = null;
                try {
                    //创建新的实体类
                    t = entity.getDeclaredConstructor().newInstance();
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                         NoSuchMethodException e) {
                    throw new RuntimeException(e);
                }
                //获取需要读取的数量
                int size = properties.size();
                for (ExcelEntity property : properties) {
                    Field field = null;
                    try {
                        //实例化字段
                        field = entity.getDeclaredField(property.getProperty());
                        field.setAccessible(true);
                        //读取当前字段在excel中的值
                        Object o = getExcelValue(getMergedRegionValue(sheet, i, property.getValue()), property);
                        //若当前字段为空，则读取数量减1
                        if (o == null || o.toString().equals("")) {
                            --size;
                        }
                        //赋值
                        field.set(t, o);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new RuntimeException(e);
                    } catch (ExcelValueError e) {
                        throw new ExcelValueError("第" + i + "行", e);
                    }
                }
                //若当前行为空行则将连续空行+1
                if (size == 0) {
                    //连续三行都是空行，则认定当前为excel结尾
                    if (++emptySize > 3) {
                        break;
                    }
                } else {
                    //若不为空行，则清空连续空行数
                    emptySize = 0;
                    varList.add(t);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return varList;
    }

    /**
     * 读取excel
     *
     * @param filepath 文件路径
     * @param filename 文件名
     * @param startrow 开始行号
     * @param startcol 开始列号
     * @param sheetnum sheet
     * @return list
     */
    public static <T> List<T> readExcel(String filepath, String filename, int startrow, int startcol, int sheetnum, Class<T> entity) {
        File target = new File(filepath, filename);
        return readExcel(target, startrow, startcol, sheetnum, entity);
    }

    public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity) {
        return readExcel(file.getFile(), startrow, startcol, sheetnum, entity);
    }

    private static Object getExcelValue(Cell cell, ExcelEntity property) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String cellValue = null;
        if (null != cell) {
            if (cell.toString().contains("-") && checkDate(cell.toString())) {
                String ans = "";
                cellValue = new SimpleDateFormat("yyyy/MM/dd").format(cell.getDateCellValue());
            } else {
                switch (cell.getCellType()) { // 判断excel单元格内容的格式，并对其进行转换，以便插入数据库
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            //判断是否为日期类型
                            cellValue = sdf.format(cell.getDateCellValue());
                        } else {
                            String msg = String.valueOf(cell.getNumericCellValue());
                            if (msg.contains(".0")) {
                                cellValue = checkNumber(String.valueOf(cell.getNumericCellValue()));
                            } else {
                                cellValue = String.valueOf(cell.getNumericCellValue());
                            }
                        }
                        break;
                    case STRING:
                        cellValue = cell.getStringCellValue();
//									 Short info=((HSSFWorkbook)cell).getCellStyle().getFont().getFontHeight();
//									cellValue=info+"";
                        break;
//								cellValue = cell.getNumericCellValue() + "";
                    case BLANK:
                        cellValue = "";
                        break;
                    case BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case ERROR:
                        cellValue = String.valueOf(cell.getErrorCellValue());
                        break;
                }
                Class<?> type = property.getType();
                ExcelFormat<?> format = property.getFormat();
                try {
                    if (format instanceof DefaultFormat) {
                        return ((DefaultFormat) format).ReadToExcel(type, cellValue);
                    }
                    return format.ReadToExcel(cellValue);
                } catch (Exception e) {
                    throw new ExcelValueError(property.getTitle() + "错误", e);
                }

            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }


    /**
     * 获取合并单元格的值
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static Cell getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    return fRow.getCell(firstColumn);
                }
            }
        }

        return sheet.getRow(row).getCell(column);
    }

    /**
     * 获取单元格的值
     *
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return cell.toString();
    }


    /**
     * 判断是否是“02-十一月-2006”格式的日期类型
     */
    private static boolean checkDate(String str) {
        String[] dataArr = str.split("-");
        try {
            if (dataArr.length == 3) {
                int x = Integer.parseInt(dataArr[0]);
                String y = dataArr[1];
                int z = Integer.parseInt(dataArr[2]);
                if (x > 0 && x < 32 && z > 0 && z < 10000 && y.endsWith("月")) {
                    return true;
                }
            }
        } catch (Exception e) {
            return false;
        }
        return false;
    }


    public static Date getDate(String time) {
        SimpleDateFormat s1 = new SimpleDateFormat("yyyy/MM/dd");
        SimpleDateFormat s2 = new SimpleDateFormat("yyyy-MM-dd");
        try {
            return s1.parse(time);
        } catch (ParseException e) {
            try {
                return s2.parse(time);
            } catch (ParseException e2) {
                e.printStackTrace();
                return null;
            }
        }
    }

    public static Date getClock(String time) {
        SimpleDateFormat s = new SimpleDateFormat("yyyy/MM/dd HH:mm");
        SimpleDateFormat s2 = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        try {
            return s.parse(time);
        } catch (ParseException e) {
            try {
                return s2.parse(time);
            } catch (ParseException e2) {
                e.printStackTrace();
                return null;
            }
        }
    }


//    public static String checkNumber(String number) {
//
//        String a = null;
//        if (number.contains(".0")) {
//            a = number.substring(0, number.length() - 2);
//        }else if(number.contains("-0")){
//            a= number;
//        } else {
//            a = number;
//        }
//        return a;
//    }

    // 工资条问题  上面的是原版的
    public static String checkNumber(String number) {

        String a = null;
        if (number.contains(".01") || number.contains(".02") || number.contains(".03") || number.contains(".04") || number.contains(".05")
                || number.contains(".06") || number.contains(".07") || number.contains(".08") || number.contains(".09")) {
            a = number;
        } else {
            if (number.contains(".0")) {
                a = number.substring(0, number.length() - 2);
            } else if (number.contains("-0")) {
                a = number;
            } else {
                a = number;
            }
        }
        return a;
    }

}
