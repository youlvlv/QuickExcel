package com.xyf.excel.util;


import com.xyf.excel.model.ExcelModel;

import javax.servlet.http.HttpServletResponse;

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
    public static DefaultDownloadExcel createDownload(HttpServletResponse response, String fileName) {
        return new DefaultDownloadExcel(response, fileName);
    }


}




