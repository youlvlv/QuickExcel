package com.lizhiwei.quickExcel.util;


import com.lizhiwei.quickExcel.model.ExcelModel;

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

    public static DefaultDownloadExcel createDownload(HttpServletResponse response, String fileName) {
        return new DefaultDownloadExcel(response, fileName);
    }


}




