package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.entity.DefaultFormat;
import com.lizhiwei.quickExcel.entity.Excel;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ExcelFormat;
import com.lizhiwei.quickExcel.exception.ExcelValueError;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
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
            //获取实体类所有属性
            Field[] fields = entity.getDeclaredFields();
            /*----------匹配头------------*/
            Row row = sheet.getRow(startrow - 1); // 行
            int cellNum = row.getLastCellNum(); // 每行的最后一个单元格位置
            for (int j = startcol; j < cellNum; j++) { // 列循环开始
//                Cell cell = row.getCell(Short.parseShort(j + ""));
//                if (cell == null) {
//                    break;
//                } else {
                ExcelEntity excelEntity = new ExcelEntity();
                //循环实体类所有属性
                for (Field field : fields) {
                    field.setAccessible(true);
                    if (!field.isAnnotationPresent(Excel.class)) {
                        continue;
                    }
                    Excel excel = field.getAnnotation(Excel.class);
                    //匹配是否为相同头
                    if (excel.value().equals(getMergedRegionValue(sheet, startrow - 1, j))) {
                        excelEntity.setProperty(field.getName());
                        excelEntity.setValue(j);
                        excelEntity.setType(field.getType());
                        try {
                            excelEntity.setFormat(excel.format().getDeclaredConstructor().newInstance());
                        } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                                 NoSuchMethodException e) {
                            e.printStackTrace();
                            excelEntity.setFormat(new DefaultFormat());
                        }
                        properties.add(excelEntity);
                        break;
                    }
                }
//                }
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
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                         NoSuchMethodException e) {
                    throw new RuntimeException(e);
                }
                int size = properties.size();
                for (ExcelEntity property : properties) {
                    Field field = null;
                    try {
                        field = entity.getDeclaredField(property.getProperty());
                        field.setAccessible(true);
                        Object o = getExcelValue(row, property);
                        if (o == null || o.toString().equals("")) {
                            --size;
                        }
                        field.set(t, o);
                    } catch (NoSuchFieldException | IllegalAccessException e) {
                        throw new RuntimeException(e);
                    } catch (ExcelValueError e) {
                        throw new ExcelValueError("第" + i + "行", e);
                    }
                }
                if (size == 0) {
                    if (++emptySize > 3) {
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
        int j = property.getValue();
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
    public static String getMergedRegionValue(Sheet sheet, int row, int column) {
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
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }

        return sheet.getRow(row).getCell(column).getStringCellValue();
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
