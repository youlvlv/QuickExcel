package com.lizhiwei.quickExcel.util;

import com.lizhiwei.quickExcel.core.ExcelUtil;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.IndexType;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.SheetModel;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.List;

public class DownloadExcel extends ExcelUtil {
    /**
     * 下载导入模板
     *
     * @param response 返回流
     * @param path     文件地址
     * @throws IOException 异常
     */
    public static void downloadTemplate(HttpServletResponse response, String path) {
        try {

            String fileName = path.substring(path.lastIndexOf("/") + 1);
            File file = new File(path);
            //作用：在前端作用显示为调用浏览器下载弹窗
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            /*response.setHeader("Content-disposition", "attachment; filename = " + new String(fileName.getBytes(fileName), "ISO8859-1"));*/
            response.setContentType("application/octet-stream");
            BufferedInputStream inputStream = new BufferedInputStream(new FileInputStream(file));
            OutputStream outputStream = response.getOutputStream();
            byte[] buffer = new byte[1024];
            int len;
            while ((len = inputStream.read(buffer)) != -1) { /** 将流中内容写出去 .*/
                outputStream.write(buffer, 0, len);
            }
            inputStream.close();
            outputStream.close();
        } catch (IOException e) {
            throw new IORunTimeException("文件操作失败", e);
        }
    }

    /**
     * 生成Excel表格以供下载
     *
     * @param listTitle
     * @param listContent
     * @param <T>
     * @Param operation 文件操作
     */
    public static <T> void setExcelProperty(FileOperation operation, List<ExcelEntity> listTitle, List<T> listContent) {
        SimpleDateFormat df = new SimpleDateFormat("MM月dd日");
        //列表排序
        try {
            //创建表格工作空间
            ExcelModel excel = new ExcelModel();
            //创建一个新表格
//            XSSFSheet xSheet = xWorkbook.createSheet(fileNameParam);
            SheetModel sheet = excel.newSheet();
            //set Sheet页头部
            util.setSheetHeader(sheet, listTitle);
            //set Sheet页内容
            util.setSheetContent(sheet, listContent, listTitle);
            excel.exportExcel(operation).close();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("导出表格时出现异常...请联系管理员", e);
        }
    }

    /**
     * 生成EXCEL表
     *
     * @param fileNameParam 文件名
     * @param response      下载流
     * @param entity        列表实体类
     * @param listContent   列表
     * @param <T>           实体类
     */
    public static <T> void setExcelProperty(String fileNameParam, HttpServletResponse response, Class<T> entity, List<T> listContent) {
        List<ExcelEntity> listTitle = util.getExcelEntities(entity);
        setExcelProperty(new DefaultDownloadExcel(response, fileNameParam), listTitle, listContent);
    }

    /**
     * 生成EXCEL表
     *
     * @param operation   文件操作
     * @param entity      列表实体类
     * @param listContent 列表
     * @param <T>         实体类
     */
    public static <T> void setExcelProperty(FileOperation operation, Class<T> entity, List<T> listContent) {
        List<ExcelEntity> listTitle = util.getExcelEntities(entity);
        setExcelProperty(operation, listTitle, listContent);
    }

    /**
     * 生成EXCEL表
     *
     * @param fileNameParam 文件名
     * @param response      下载流
     * @param entity        列表实体类
     * @param listContent   列表
     * @param <T>           实体类
     */
    public static <T> void setExcelProperty(String fileNameParam, HttpServletResponse response, Class<T> entity, List<T> listContent, IndexType type) {
        List<ExcelEntity> listTitle = util.getExcelEntities(entity, true, type);
        setExcelProperty(new DefaultDownloadExcel(response, fileNameParam), listTitle, listContent);
    }

    private DownloadExcel() {
    }
}
