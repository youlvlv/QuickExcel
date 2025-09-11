package com.lizhiwei.quickExcel.model;

import java.io.IOException;
import java.io.OutputStream;

public interface HttpServletResponseModel {


    OutputStream apply(String fileName) throws IOException;

}
