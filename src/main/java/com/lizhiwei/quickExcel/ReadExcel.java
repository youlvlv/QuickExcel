package com.lizhiwei.quickExcel;


import com.lizhiwei.quickExcel.entity.Excel;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ReadExcel {


    /**
     * @param filepath //文件路径
     * @param filename //文件名
     * @param startrow //开始行号
     * @param startcol //开始列号
     * @param sheetnum //sheet
     * @return list
     */
    public static <T> List<T> readExcel(String filepath, String filename, int startrow, int startcol, int sheetnum, Class<T> entity) {
        List<T> varList = new ArrayList<>();

        try {
//			File target = new File(filepath, filename);
//			FileInputStream fi = new FileInputStream(target);
//
//			HSSFWorkbook wb = new HSSFWorkbook(fi);
//			HSSFSheet sheet = wb.getSheetAt(sheetnum); // sheet 从0开始
//			int rowNum = sheet.getLastRowNum() + 1; // 取得最后一行的行号
            File target = new File(filepath, filename);
            FileInputStream fi = new FileInputStream(target);
            String fileType = filename.substring(filename.lastIndexOf(".") + 1);
            Workbook wb = null;
            if (fileType.equals("xls")) {
                wb = new HSSFWorkbook(fi);
            } else if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(fi);
            }
            Sheet sheet = wb.getSheetAt(sheetnum); // sheet 从0开始
            List<ExcelEntity> properties = new ArrayList<>();
            Field[] fields = entity.getDeclaredFields();
            /*----------匹配头------------*/
            Row row = sheet.getRow(startrow - 1); // 行
            int cellNum = row.getLastCellNum(); // 每行的最后一个单元格位置
            for (int j = startcol; j < cellNum; j++) { // 列循环开始
                Cell cell = row.getCell(Short.parseShort(j + ""));
                if (cell == null) {
                    break;
                } else {
                    ExcelEntity excelEntity = new ExcelEntity();
                    for (Field field : fields) {
                        field.setAccessible(true);
                        if (!field.isAnnotationPresent(Excel.class)) {
                            continue;
                        }
                        Excel excel = field.getAnnotation(Excel.class);
                        if (excel.value().equals(cell.getStringCellValue())) {
                            excelEntity.setProperty(field.getName());
                            excelEntity.setValue(j + "");
                            excelEntity.setType(field.getGenericType().getTypeName());
                            properties.add(excelEntity);
                            break;
                        }
                    }
                }
            }
            int rowNum = sheet.getLastRowNum() + 1; // 取得最后一行的行号
            //空行数
            int emptySize = 0;
            for (int i = startrow; i < rowNum; i++) { // 行循环开始

//				PageData varpd = new PageData();
//				HSSFRow row = sheet.getRow(i); // 行
//				int cellNum = row.getLastCellNum(); // 每行的最后一个单元格位置
                row = sheet.getRow(i); // 行
                if (row == null) {
                    break;
                }
                T t = null;
                try {
                    //创建新的实体类
                    t = entity.getDeclaredConstructor().newInstance();
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                    throw new RuntimeException(e);
                }
                int size = properties.size();
                for (ExcelEntity property : properties) {
                    Field field = null;
                    try {
                        field = entity.getDeclaredField(property.getProperty());
                        field.setAccessible(true);
                        Object o = getExcelValue(row, property);
                        if (o == null){
                            --size;
                        }
                        field.set(t, o);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new RuntimeException(e);
                    }
                }
                if (size == 0) {
                    if (++emptySize > 3){
                        break;
                    }
                } else {
                    emptySize = 0;
                    varList.add(t);
                }

            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return varList;
    }

    private static Object getExcelValue(Row row, ExcelEntity property) {
        int j = Integer.parseInt(property.getValue());
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        Cell cell = row.getCell(Short.parseShort(j + ""));
        String cellValue = null;
        if (null != cell) {
            if (cell.toString().contains("-") && checkDate(cell.toString())) {
                String ans = "";
                    cellValue = new SimpleDateFormat("yyyy/MM/dd").format(cell.getDateCellValue());
            } else {
                switch (cell.getCellTypeEnum()) { // 判断excel单元格内容的格式，并对其进行转换，以便插入数据库
                    case NUMERIC:
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
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
                Class type = null;
                try {
                    type = Class.forName(property.getType());
                } catch (ClassNotFoundException e) {
                    throw new RuntimeException(e);
                }
                if (cellValue == null) {
                    return null;
                } else {
                    if (type == String.class) {
                        return cellValue;
                    } else if (type == Integer.class) {
                        return Integer.parseInt(cellValue);
                    } else if (type == Long.class) {
                        return Long.valueOf(cellValue);
                    } else if (type == Double.class) {
                        return Double.valueOf(cellValue);
                    } else if (type == Boolean.class) {
                        return Boolean.valueOf(cellValue);
                    } else if (type == Date.class) {
                        try {
                            return sdf.parse(cellValue);
                        } catch (ParseException e) {
                            throw new RuntimeException(e);
                        }
                    }
                }

            }
        } else {
            cellValue = "";
        }
        return cellValue;
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

    public static String checkNumber(String number) {
        String a = null;
        if (number.contains(".0")) {
            a = number.substring(0, number.length() - 2);
        } else {
            a = number;
        }
        return a;
    }

}
