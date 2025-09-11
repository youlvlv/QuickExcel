package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.HttpServletResponseJavaxModel;
import com.lizhiwei.quickExcel.model.HttpServletResponseSpring6Model;


/**
 * 通过链式方式构建Excel
 * 可用于多sheet构建
 */
public class DownloadComplexExcel {
    /**
     * 创建新的excel
     *
     * @return
     */
    public static ExcelModel newExcel() {
        return new ExcelModel();
    }

    /**
     * 创建默认的下载excel工具
     * @param response
     * @param fileName
     * @return
     */
    public static DefaultDownloadExcel createDownload(Object response, String fileName) {
        try {
            Class.forName("jakarta.servlet.http.HttpServletResponse");
            return new DefaultDownloadExcel(HttpServletResponseSpring6Model.of(response), fileName);
        } catch (ClassNotFoundException e) {
            return new DefaultDownloadExcel(HttpServletResponseJavaxModel.of(response), fileName);
        }

    }


}




