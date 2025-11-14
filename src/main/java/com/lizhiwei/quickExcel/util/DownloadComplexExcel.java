package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.HttpServletResponseModel;


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
        HttpServletResponseModel model;
        try {
            Class.forName("jakarta.servlet.http.HttpServletResponse");
            model = ConditionalProxyFactory.createProxyIfAvailable(
                    Class.forName("com.lizhiwei.quickExcel.model.HttpServletResponseSpring6Model$HttpServletResponseSpring6ModelFactory"),
                    HttpServletResponseModel.HttpServletResponseModelFactory.class).of(response);
            return new DefaultDownloadExcel(model, fileName);
        } catch (ClassNotFoundException e) {
            try {
                model = ConditionalProxyFactory.createProxyIfAvailable(
                        Class.forName("com.lizhiwei.quickExcel.model.HttpServletResponseJavaxModel$HttpServletResponseJavaxModelFactory"),
                        HttpServletResponseModel.HttpServletResponseModelFactory.class).of(response);
                return new DefaultDownloadExcel(model, fileName);
            } catch (ClassNotFoundException ex) {
                throw new RuntimeException(ex);
            }

        }

    }


}




