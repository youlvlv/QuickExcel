package com.lizhiwei.quickExcel.v3.read.reader;

import com.lizhiwei.quickExcel.v3.read.model.ImageData;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelReader {
    
    <T> List<T> readExcel(File file, int startRow, int sheetNum, Class<T> entity);
    
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity);
    
    <T> List<T> readExcel(File file, int startRow, int startCol, int sheetNum, Class<T> entity, boolean safe);
    
    List<Map<String, String>> readExcelAsMap(File file, int startRow, int startCol, int sheetNum, List<String> headers);
    
    List<ImageData> readExcelImages(File file, int sheetNum);
}
