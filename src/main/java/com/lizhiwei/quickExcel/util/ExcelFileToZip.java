package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.HttpServletResponseModel;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 批量导出excel到zip压缩包
 */
public class ExcelFileToZip {

    private List<File> files = new ArrayList<>();

    private File path = new File("cache");

    /*
     *实例化后执行
     */ {
        if (path.exists()) {
            path.mkdir();
        }
    }


    /**
     * 设置当前excel文件名称
     *
     * @param fileName
     * @return
     */
    public FileOperation fileName(String fileName) {
        return (operation) -> {
            File file = new File(path, fileName + ".xlsx");
            try (FileOutputStream outFile = new FileOutputStream(file)) {
                operation.accept(outFile);
            } catch (IOException e) {
                throw new RuntimeException(e);
            } finally {
                files.add(file);
            }
        };
    }


    /**
     * 下载zip文件
     *
     * @param response
     * @param name
     */
    public void downloadZip(HttpServletResponseModel response, String name) {
        name = name + ".zip";
        File zip = new File(path, name);
        toZip(files, zip);
        try (OutputStream outputStream = response.getOutputStream(name)) {
            BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(zip.toPath()));
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
            files.forEach(File::delete);
            zip.delete();
        }
    }

    /**
     * 将文件添加进压缩包
     *
     * @param srcFiles 源文件
     * @param zipFile  压缩包文件
     * @throws RuntimeException
     */
    public void toZip(List<File> srcFiles, File zipFile) throws RuntimeException {
        //判断压缩文件是否为空
        if (zipFile == null) {
            return;
        }
        //判断是否为zip压缩文件
        if (!zipFile.getName().endsWith(".zip")) {
            return;
        }
        ZipOutputStream zos = null;
        try {
            FileOutputStream out = new FileOutputStream(zipFile);
            zos = new ZipOutputStream(out);
            for (File srcFile : srcFiles) {
                byte[] buf = new byte[4096];
                zos.putNextEntry(new ZipEntry(srcFile.getName()));
                int len;
                FileInputStream in = new FileInputStream(srcFile);
                while ((len = in.read(buf)) != -1) {
                    zos.write(buf, 0, len);
                }
                in.close();
                zos.closeEntry();
            }
            zos.close();
            out.close();
        } catch (IOException e) {
            throw new IORunTimeException(e);
        }
    }
}
