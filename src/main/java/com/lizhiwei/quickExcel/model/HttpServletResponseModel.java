package com.lizhiwei.quickExcel.model;

import java.io.IOException;
import java.io.OutputStream;

public interface HttpServletResponseModel {


    OutputStream getOutputStream(String fileName) throws IOException;

    interface HttpServletResponseModelFactory {
        HttpServletResponseModel of(Object response);
    }

}
