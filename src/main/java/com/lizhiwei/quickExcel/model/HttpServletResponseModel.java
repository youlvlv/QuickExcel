package com.lizhiwei.quickExcel.model;

import java.io.IOException;
import java.io.OutputStream;
import java.util.function.Function;

public interface HttpServletResponseModel {


    OutputStream apply(String fileName) throws IOException;

}
