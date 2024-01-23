package com.xyf.excel.util;

import com.xyf.excel.exception.IORunTimeException;
import com.xyf.excel.model.ExcelModel;
import com.xyf.excel.model.FileOperation;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 默认下载excel工具类
 */
class DefaultDownloadExcel implements FileOperation {

    private final HttpServletResponse response;

    private final String fileNameParam;

    private final SimpleDateFormat df = new SimpleDateFormat("MM月dd日");

    /**
     * 开始下载
     *
     * @param excel
     */
    public void download(ExcelModel excel) {
        if (!new File("cache").exists()) {
            if (!new File("cache").mkdir()) {
                throw new IORunTimeException("无法创建缓存文件夹");
            }
        }
        String fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
        String fileName2 = "cache/" + fileName;
        try( FileOutputStream outFile = new FileOutputStream(fileName2);) {
            excel.write(outFile);
        } catch (IOException e) {
            throw new IORunTimeException(e);
        }
        downloadTemplate(response, fileName2);
    }

    private static void downloadTemplate(HttpServletResponse response, String path) {
        String fileName = path.substring(path.lastIndexOf("/") + 1);
        File file = new File(path);
        try( BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(file.toPath()));
             OutputStream outputStream = response.getOutputStream();) {
            //作用：在前端作用显示为调用浏览器下载弹窗
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setContentType("application/octet-stream");

            byte[] buffer = new byte[1024];
            int len;
            while ((len = inputStream.read(buffer)) != -1) { /** 将流中内容写出去 .*/
                outputStream.write(buffer, 0, len);
            }
        } catch (IOException e) {
            throw new IORunTimeException("文件操作失败", e);
        } finally {
            file.delete();
        }
    }

    DefaultDownloadExcel(HttpServletResponse response, String fileName) {
        this.response = response;
        this.fileNameParam = fileName;
    }

    @Override
    public void run(ExcelModel model) {
        this.download(model);
    }
}
