package com.lizhiwei.quickExcel.util;

import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Date;

class DefaultDownloadExcel implements FileOperation {

    private HttpServletResponse response;

    private String fileNameParam;

    private final SimpleDateFormat df = new SimpleDateFormat("MM月dd日");

    public void download(ExcelModel excel) throws FileNotFoundException {
        String fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
        String fileName2 = "cache/" + fileName;
        FileOutputStream outFile = new FileOutputStream(fileName2);
        try {
            excel.write(outFile);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        downloadTemplate(response, fileName2);
    }

    public static void downloadTemplate(HttpServletResponse response, String path) {
        try {

            String fileName = path.substring(path.lastIndexOf("/") + 1);
            File file = new File(path);
            //作用：在前端作用显示为调用浏览器下载弹窗
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            /*response.setHeader("Content-disposition", "attachment; filename = " + new String(fileName.getBytes(fileName), "ISO8859-1"));*/
            response.setContentType("application/octet-stream");
            BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(file.toPath()));
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
        } finally {
            File file = new File(path);
            file.delete();
        }
    }

    DefaultDownloadExcel(HttpServletResponse response, String fileName) {
        this.response = response;
        this.fileNameParam = fileName;
    }

    @Override
    public void run(ExcelModel model) {
        try {
            this.download(model);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }
}
