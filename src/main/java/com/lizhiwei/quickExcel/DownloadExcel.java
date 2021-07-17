package com.lizhiwei.quickExcel;


import com.lizhiwei.quickExcel.entity.Excel;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ExcelFormat;



import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class DownloadExcel {
    /**
     * 下载导入模板
     *
     * @param response 返回流
     * @param path     文件地址
     * @throws IOException 异常
     */
//    public static void download(HttpServletResponse response, File path) throws IOException {
//            String fileName = "文件名";
//            File file = this.setExcelProperty()
//            /** 将文件名称进行编码 */
//            response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
//            response.setContentType("application/octet-stream");
//            BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(file));
//            OutputStream outputStream = response.getOutputStream();
//            byte[] buffer = new byte[1024];
//            int len;
//            while ((len = inputStream.read(buffer)) != -1) { /** 将流中内容写出去 .*/
//                outputStream.write(buffer, 0, len);
//            }
//            inputStream.close();
//            outputStream.close();
//    }

    /**
     * 生成Excel表格以供下载
     *
     * @param fileNameParam
     * @param listTitle
     * @param listContent
     * @param <T>
     */
    public static <T> File setExcelProperty(String fileNameParam,  List<ExcelEntity> listTitle, List<T> listContent) {
        SimpleDateFormat df = new SimpleDateFormat("MM月dd日");
        XSSFWorkbook xWorkbook = null;
        String fileName = "";
        String fileName2 = "";
        try {
            //定义表格导出时默认文件名 时间戳
            //String fileName = df.format(new Date()) + ".xlsx";
            fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
            fileName2 = "cache/" + fileName;
            //创建表格工作空间
            xWorkbook = new XSSFWorkbook();
            //创建一个新表格
            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
            //set Sheet页头部
            setSheetHeader(xWorkbook, xSheet, listTitle);
            //set Sheet页内容
            setSheetContent(xWorkbook, xSheet, listContent, listTitle);
            FileOutputStream outFile = new FileOutputStream(fileName2);
            xWorkbook.write(outFile);
            xWorkbook.close();


        } catch (Exception e) {
            e.printStackTrace();
        }
        return new File(fileName2);
    }

    /**
     * 生成EXCEL表
     * @param fileNameParam 文件名
     * @param entity 列表实体类
     * @param listContent 列表
     * @param <T> 实体类
     */
    public static <T> File setExcelProperty(String fileNameParam, Class<T> entity, List<T> listContent) {
        SimpleDateFormat df = new SimpleDateFormat("MM月dd日");
        XSSFWorkbook xWorkbook = null;
        String fileName = "";
        String fileName2 = "";
        Field[] fields = entity.getDeclaredFields();
        List<ExcelEntity> listTitle = new ArrayList<>();
        for (Field field : fields) {
            //设置属性默认可访问，防止private阻止访问
            field.setAccessible(true);
            //判断是否包含Excel注解
            if (field.isAnnotationPresent(Excel.class)) {
                //获取Excel注解
                Excel e = field.getDeclaredAnnotation(Excel.class);
                ExcelEntity excelEntity = new ExcelEntity(field.getName(), e.name(),e.format());
                listTitle.add(excelEntity);
            }
        }
        try {
            //定义表格导出时默认文件名 时间戳
            //String fileName = df.format(new Date()) + ".xlsx";
            fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
            fileName2 = "cache/" + fileName;
            //创建表格工作空间
            xWorkbook = new XSSFWorkbook();
            //创建一个新表格
            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
            //set Sheet页头部
            setSheetHeader(xWorkbook, xSheet, listTitle);
            //set Sheet页内容
            setSheetContent(xWorkbook, xSheet, listContent, listTitle);
            //判断路径是否存在
            File file=new File(fileName2);
            if (!file.getParentFile().exists()) {
                //若不存在则新建
                file.getParentFile().mkdirs();
            }
            FileOutputStream outFile = new FileOutputStream(file);
            xWorkbook.write(outFile);
            xWorkbook.close();


        } catch (Exception e) {
            e.printStackTrace();
            //更换为自定异常！！！！
            throw new RuntimeException("导出表格时出现异常...请联系管理员",e);
        }
        return new File(fileName2);
    }


    /**
     * 配置Excel表格的顶部信息，如：学号  姓名  年龄  出生年月
     *
     * @param xWorkbook
     * @param xSheet
     */
    private static void setSheetHeader(XSSFWorkbook xWorkbook, XSSFSheet xSheet, List<ExcelEntity> listTitle) {
        //设置表格的宽度  xSheet.setColumnWidth(0, 20 * 256); 中的数字 20 自行设置为自己适用的
        /*xSheet.setColumnWidth(0, 20 * 256);
        xSheet.setColumnWidth(1, 15 * 256);
        xSheet.setColumnWidth(2, 15 * 256);
        xSheet.setColumnWidth(3, 20 * 256);*/

        //创建表格的样式
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

        //创建一行
        XSSFRow xRow0 = xSheet.createRow(0);
        int i = 0;
        for (ExcelEntity excelEntity : listTitle) {
            XSSFCell xCell0 = xRow0.createCell(i);
            xCell0.setCellStyle(cs);
            xCell0.setCellValue(excelEntity.getTitle());
            i++;
        }
    }

    /**
     * 配置(赋值)表格内容部分
     *
     * @param xWorkbook
     * @param xSheet
     * @param listContent
     * @throws Exception
     */
    private static <T> void setSheetContent(XSSFWorkbook xWorkbook, XSSFSheet xSheet, List<T> listContent, List<ExcelEntity> listTitle) throws Exception {

        //创建内容样式（头部以下的样式）
        CellStyle cs = xWorkbook.createCellStyle();
        cs.setWrapText(true);

        //设置水平垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        if (null != listContent && listContent.size() > 0) {
            try {
                for (int i = 0; i < listContent.size(); i++) {
                    XSSFRow xRow = xSheet.createRow(i + 1);
                    //获取类属性
                    Field field;
                    int j = 0;
                    for (ExcelEntity excelEntity : listTitle) {
                        String str = excelEntity.getValue();
                        //获取完成get方法  首字母大写如：getId
                        field = listContent.get(i).getClass().getDeclaredField(str);
                        field.setAccessible(true);
                        Object o = field.get(listContent.get(i));
                        String value = "";
                        ExcelFormat format = excelEntity.getFormat().getConstructor().newInstance();
                        value = format.WriterToExcel(o).toString();
                        //循环设置每列的值
                        XSSFCell xCell = xRow.createCell(j);
                        xCell.setCellStyle(cs);
                        xCell.setCellValue(value);
                        j++;
                    }
                }
            } catch (IllegalAccessException e) {
                System.out.println(e.getMessage());
            }
        }

    }


    /**
     * 生成Excel表格以供下载
     *
     * @param fileNameParam
     * @param response
     * @param listTitle
     * @param listContent
     * @param
     */
    public static <T,E> File setExcelProperty(String fileNameParam, List<Map<T,E>> listTitle, List<Map<T,E>> listContent, Integer z) {
        SimpleDateFormat df = new SimpleDateFormat("MM月dd日");
        XSSFWorkbook xWorkbook = null;
        String fileName = "";
        String fileName2 = "";
        try {
            //定义表格导出时默认文件名 时间戳
            //String fileName = df.format(new Date()) + ".xlsx";
            fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
            fileName2 = "cache/" + fileName;

            //创建表格工作空间
            xWorkbook = new XSSFWorkbook();
            //创建一个新表格
            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
            //set Sheet页头部
            setSheetHeader(xWorkbook, xSheet, listTitle, 1);
            //set Sheet页内容
            setSheetContent(xWorkbook, xSheet, listContent, listTitle, 1);
            FileOutputStream outFile = new FileOutputStream(fileName2);
            xWorkbook.write(outFile);
            xWorkbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
        return  new File(fileName2);
    }


    /**
     * 配置Excel表格的顶部信息，如：学号  姓名  年龄  出生年月
     *
     * @param xWorkbook
     * @param xSheet
     */
    private static <T,E> void setSheetHeader(XSSFWorkbook xWorkbook, XSSFSheet xSheet, List<Map<T,E>> listTitle, Integer z) {
        //设置表格的宽度  xSheet.setColumnWidth(0, 20 * 256); 中的数字 20 自行设置为自己适用的
        /*xSheet.setColumnWidth(0, 20 * 256);
        xSheet.setColumnWidth(1, 15 * 256);
        xSheet.setColumnWidth(2, 15 * 256);
        xSheet.setColumnWidth(3, 20 * 256);*/

        //创建表格的样式
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

        //创建一行
        XSSFRow xRow0 = xSheet.createRow(0);
        int i = 0;
        for (Map excelEntity : listTitle) {
            XSSFCell xCell0 = xRow0.createCell(i);
            xCell0.setCellStyle(cs);
            xCell0.setCellValue(excelEntity.get("title").toString());
            i++;
        }
    }

    /**
     * 配置(赋值)表格内容部分
     *
     * @param xWorkbook
     * @param xSheet
     * @param listContent
     * @throws Exception
     */
    private static <T,E> void setSheetContent(XSSFWorkbook xWorkbook, XSSFSheet xSheet, List<Map<T,E>> listContent, List<Map<T,E>> listTitle, Integer z) throws Exception {

        //创建内容样式（头部以下的样式）
        CellStyle cs = xWorkbook.createCellStyle();
        cs.setWrapText(true);

        //设置水平垂直居中
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        if (null != listContent && listContent.size() > 0) {
            for (int i = 0; i < listContent.size(); i++) {
                XSSFRow xRow = xSheet.createRow(i + 1);
                //获取类属性
                int j = 0;
                for (Map excelEntity : listTitle) {
                    String str = excelEntity.get("value").toString();
                    //获取完成get方法  首字母大写如：getId
                    String value = listContent.get(i).get(str).toString();
                    //循环设置每列的值
                    XSSFCell xCell = xRow.createCell(j);
                    xCell.setCellStyle(cs);
                    xCell.setCellValue(value);
                    j++;
                }
            }
        }

    }


}
