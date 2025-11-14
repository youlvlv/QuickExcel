package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.exception.IORunTimeException;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

public class HttpServletResponseJavaxModel implements HttpServletResponseModel {

    private final HttpServletResponse response;

    public HttpServletResponseJavaxModel(HttpServletResponse response) {
        this.response = response;
    }

    public static HttpServletResponseModel of(Object response) {
        if (response instanceof HttpServletResponse response1) {
            return new HttpServletResponseJavaxModel(response1);
        }
        throw new IORunTimeException("当前支持的为 Spring5 的 javax.servlet.http.HttpServletResponse,请检查您的传入参数");
    }

    @Override
    public OutputStream getOutputStream(String fileName) throws IOException {
        response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8));
        response.setContentType("application/octet-stream");
        return response.getOutputStream();
    }

    public static class HttpServletResponseJavaxModelFactory implements HttpServletResponseModelFactory {
        private static HttpServletResponseModel create(Object response) {
            return HttpServletResponseJavaxModel.of(response);
        }

        @Override
        public HttpServletResponseModel of(Object response) {
            return create(response);
        }
    }
}
