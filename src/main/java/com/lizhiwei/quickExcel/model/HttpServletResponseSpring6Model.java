package com.lizhiwei.quickExcel.model;

import com.lizhiwei.quickExcel.exception.IORunTimeException;
import jakarta.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

public class HttpServletResponseSpring6Model implements HttpServletResponseModel {

    private final HttpServletResponse response;

    public HttpServletResponseSpring6Model(HttpServletResponse response) {
        this.response = response;
    }

    @Override
    public OutputStream getOutputStream(String fileName) throws IOException {
        response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, StandardCharsets.UTF_8));
        response.setContentType("application/octet-stream");
        return response.getOutputStream();

    }

    public static HttpServletResponseModel of(Object response) {
        if (response instanceof HttpServletResponse response1) {
            return new HttpServletResponseSpring6Model(response1);
        }
        throw new IORunTimeException("当前支持的为 Spring6 的 jakarta.servlet.http.HttpServletResponse,请检查您的传入参数");
    }


    public static class HttpServletResponseSpring6ModelFactory implements HttpServletResponseModelFactory {
        private static HttpServletResponseModel create(Object response) {
            return HttpServletResponseSpring6Model.of(response);
        }

        @Override
        public HttpServletResponseModel of(Object response) {
            return create(response);
        }
    }
}
