package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.exception.IORunTimeException;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;

public class HttpServletResponseSpring6Model implements HttpServletResponseModel {

    private final HttpServletResponse response;

    public HttpServletResponseSpring6Model(HttpServletResponse response) {
        this.response = response;
    }

    @Override
    public OutputStream apply(String fileName) throws IOException {
        response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        response.setContentType("application/octet-stream");
        return response.getOutputStream();

    }

    public static HttpServletResponseModel of(Object response) {
        if (response instanceof HttpServletResponse response1) {
            return new HttpServletResponseSpring6Model(response1);
        }
        throw new IORunTimeException("当前支持的为 Spring6 的 jakarta.servlet.http.HttpServletResponse,请检查您的传入参数");
    }
}
