package com.lizhiwei.quickExcel.util;

import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.HttpServletResponseModel;
import jakarta.servlet.http.HttpServletResponse;

import java.io.*;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.function.Consumer;

/**
 * 默认下载excel工具类
 */
class DefaultDownloadExcel implements FileOperation {

	private final HttpServletResponseModel response;

	private final String fileNameParam;

	private final SimpleDateFormat df = new SimpleDateFormat("MM月dd日");

	DefaultDownloadExcel(HttpServletResponseModel response, String fileName) {
		this.response = response;
		this.fileNameParam = fileName;
	}

	private void downloadTemplate( String path) {
		String fileName = path.substring(path.lastIndexOf("/") + 1);
		File file = new File(path);
		try (BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(file.toPath()));
		     OutputStream outputStream = response.apply(fileName);) {
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

	/**
	 * 开始下载
	 *
	 * @param excel
	 */
	public void download(Consumer<OutputStream> excel) {
		if (!new File("cache").exists()) {
			if (!new File("cache").mkdir()) {
				throw new IORunTimeException("无法创建缓存文件夹");
			}
		}
		String fileName = df.format(new Date()) + "-" + fileNameParam + ".xlsx";
		String fileName2 = "cache/" + fileName;
		try (FileOutputStream outFile = new FileOutputStream(fileName2);) {
			excel.accept(outFile);
		} catch (IOException e) {
			throw new IORunTimeException(e);
		}
		downloadTemplate(fileName2);
	}

	@Override
	public void run(Consumer<OutputStream> model) {
		this.download(model);
	}
}
