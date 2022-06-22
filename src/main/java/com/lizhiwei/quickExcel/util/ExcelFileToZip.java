package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;


import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ExcelFileToZip implements FileOperation {

    private List<File> files = new ArrayList<>();

    private String fileName;

    private File path = new File("cache");

    /*
      实例化后执行
     */
    {
        if (path.exists()) {
            path.mkdir();
        }
    }

    @Override
    public void run(ExcelModel model) {
        File file = new File(path, fileName + ".xlsx");
        try {
            FileOutputStream outFile = new FileOutputStream(file);
            model.getWorkbook().write(outFile);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        files.add(file);
    }

    /**
     * 设置当前excel文件名称
     *
     * @param fileName
     * @return
     */
    public ExcelFileToZip fileName(String fileName) {
        this.fileName = fileName;
        return this;
    }

    /**
     * 下载zip文件
     *
     * @param response
     * @param name
     */
    public void downloadZip(HttpServletResponse response, String name) {
        name = name + ".zip";
        File zip = new File(path, name);
        toZip(files, zip);
        try {
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(name, "UTF-8"));
            response.setContentType("application/octet-stream");
            BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(zip.toPath()));
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
    public static void toZip(List<File> srcFiles, File zipFile) throws RuntimeException {
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
                System.out.println("srcFile.getName()" + srcFile.getName());
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
