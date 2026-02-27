package com.lizhiwei.quickExcel.v2;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.PictureMap;
import com.lizhiwei.quickExcel.entity.ReadErrorInfo;
import com.lizhiwei.quickExcel.entity.Rule;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.exception.ExcelValueException;
import com.lizhiwei.quickExcel.exception.IORunTimeException;
import com.lizhiwei.quickExcel.format.DefaultFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;
import com.lizhiwei.quickExcel.model.ExcelBaseModel;
import com.lizhiwei.quickExcel.model.ExcelModel;
import com.lizhiwei.quickExcel.model.UploadFile;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class ReadExcelByPoi {

    private static final Map<String, String> MIME_TYPE_TO_EXTENSION = new HashMap<>();
    static {
        MIME_TYPE_TO_EXTENSION.put("image/jpeg", ".jpg");
        MIME_TYPE_TO_EXTENSION.put("image/png", ".png");
        MIME_TYPE_TO_EXTENSION.put("image/gif", ".gif");
        MIME_TYPE_TO_EXTENSION.put("image/bmp", ".bmp");
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity) {
        return readExcel(file, startrow, startcol, sheetnum, entity, false, false);
    }

    public static ExcelModel readExcel(File file) {
        try {
            return new ExcelModel(getWorkbook(file));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, String sheetName, Class<T> entity,
                                        boolean safe, boolean readImage) {
        int sheetnum = 0;
        try {
            Workbook wb = getWorkbook(file);
            sheetnum = wb.getSheetIndex(wb.getSheet(sheetName));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return readExcel(file, startrow, startcol, sheetnum, entity, safe, readImage);
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, String sheetName, Class<T> entity,
                                        boolean safe) {
        int sheetnum = 0;
        try {
            Workbook wb = getWorkbook(file);
            sheetnum = wb.getSheetIndex(wb.getSheet(sheetName));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return readExcel(file, startrow, startcol, sheetnum, entity, safe, false);
    }

    public static List<Map<String, String>> readExcelAsMap(File file, int startrow, int startcol, int sheetnum, 
                                                           boolean safe, List<ExcelEntity> propertieList) {
        List<Map<String, String>> varList = new ArrayList<>();
        boolean error = false;
        List<ReadErrorInfo> errorInfoList = new ArrayList<>();
        try {
            Workbook wb = getWorkbook(file);
            Sheet sheet = wb.getSheetAt(sheetnum);
            List<ExcelEntity> properties = getExcelEntities(startrow, startcol, propertieList, sheet);
            Row row;
            int rowNum = sheet.getLastRowNum() + 1;
            int emptySize = 0;
            for (int i = startrow; i < rowNum; i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                Map<String, String> t = new HashMap<>();
                int size = properties.size();
                for (ExcelEntity property : properties) {
                    Field field = null;
                    Method method = null;
                    try {
                        String o = getExcelStringValue(wb, getMergedRegionValue(sheet, i, property.getValue()), property);
                        if (o == null || o.isEmpty()) {
                            --size;
                        }
                        t.put(property.getProperty(), o);
                    } catch (ExcelValueException e) {
                        if (safe) {
                            error = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + i + "行" + " " + e.getMessage(), e);
                        }
                    }
                }
                if (size == 0) {
                    if (++emptySize > 3) {
                        break;
                    }
                } else {
                    emptySize = 0;
                    varList.add(t);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        if (error) {
            throw new ExcelReadException(errorInfoList);
        }
        if (varList.isEmpty()) {
            throw new ExcelReadException("当前表格为空");
        }
        return varList;
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity,
                                        boolean safe, boolean readImage, List<ExcelEntity> propertieList) {
        List<T> varList = new ArrayList<>();
        boolean error = false;
        List<ReadErrorInfo> errorInfoList = new ArrayList<>();
        try {
            Workbook wb = getWorkbook(file);
            Sheet sheet = wb.getSheetAt(sheetnum);
            PictureMap pictureMap = new PictureMap();
            if (readImage) {
                Drawing<?> drawing = sheet.getDrawingPatriarch();
                Optional.ofNullable(drawing).ifPresent(draw -> {
                    for (Object o : draw) {
                        if (o instanceof Picture picture) {
                            ClientAnchor ca = picture.getClientAnchor();
                            pictureMap.put(new PictureMap.PictureKey(ca.getRow1(), (int) ca.getCol1()), picture);
                        }
                    }
                });
            }
            List<ExcelEntity> properties = getExcelEntities(startrow, startcol, propertieList, sheet);
            Row row;
            int rowNum = sheet.getLastRowNum() + 1;
            int emptySize = 0;
            for (int i = startrow; i < rowNum; i++) {
                row = sheet.getRow(i);
                if (row == null) {
                    break;
                }
                T t = null;
                try {
                    t = entity.getDeclaredConstructor().newInstance();
                } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                         NoSuchMethodException e) {
                    throw new RuntimeException("构建实体类失败！请检查实体类", e);
                }
                int pictureIndex = 0;
                int size = properties.size();
                Map<String, String> objectMap = new HashMap<>();
                Map<String, ExcelEntity> propertyMap = new HashMap<>();
                for (ExcelEntity property : properties) {
                    try {
                        Cell cell = getMergedRegionValue(sheet, i, property.getValue());
                        String o = getCellStringValue(wb, cell, property);
                        objectMap.put(property.getProperty(), o);
                        if (!Objects.equals(property.getAliasProperty(), "")) {
                            if (!property.getAliasProperty().equals(property.getProperty())) {
                                if (objectMap.containsKey(property.getAliasProperty())) {
                                    throw new ExcelReadException("存在了重复的 property");
                                }
                                objectMap.put(property.getAliasProperty(), o);
                            }
                        }
                        propertyMap.put(property.getProperty(), property);
                        if (o.isEmpty()) {
                            --size;
                        }
                    } catch (ExcelValueException e) {
                        if (safe) {
                            error = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + i + "行" + " " + e.getMessage(), e);
                        }
                    }
                }
                for (Map.Entry<String, ExcelEntity> entry : propertyMap.entrySet()) {
                    try {
                        ExcelEntity property = entry.getValue();
                        Object o = getCellValue(objectMap.get(entry.getKey()), property, objectMap);
                        for (Rule rule : property.getRules()) {
                            rule.rule(o);
                        }
                        Field field;
                        Method method;
                        switch (property.getParamType()) {
                            case FIELD: {
                                field = entity.getDeclaredField(property.getProperty());
                                field.setAccessible(true);
                                field.set(t, o);
                                break;
                            }
                            case METHOD: {
                                String set = "set" + Pattern.compile("^.").matcher(property.getProperty()).replaceFirst(m -> m.group().toUpperCase());
                                method = entity.getMethod(set, property.getType());
                                method.invoke(t, o);
                                break;
                            }
                            case IMAGE: {
                                Picture picture = pictureMap.get(new PictureMap.PictureKey(row.getRowNum(), property.getValue()));
                                if (picture == null) {
                                    break;
                                }
                                field = entity.getDeclaredField(property.getProperty());
                                field.setAccessible(true);
                                field.set(t, formatValue(property,
                                        ExcelConfig.getImageFileFunction().apply(new ByteArrayInputStream(picture.getPictureData().getData()), getPictureExtension(picture))));
                                break;
                            }
                        }
                    } catch (NoSuchFieldException | IllegalAccessException | NoSuchMethodException |
                             InvocationTargetException e) {
                        throw new RuntimeException(e);
                    } catch (ExcelValueException e) {
                        if (safe) {
                            error = true;
                            errorInfoList.add(new ReadErrorInfo(i, e.getMessage()));
                        } else {
                            throw new ExcelReadException("第" + (i + startrow) + "行" + " " + e.getMessage(), e);
                        }
                    }
                }
                if (size == 0) {
                    if (++emptySize > 3) {
                        break;
                    }
                } else {
                    emptySize = 0;
                    varList.add(t);
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        if (error) {
            throw new ExcelReadException(errorInfoList);
        }
        if (varList.isEmpty()) {
            throw new ExcelReadException("当前表格为空");
        }
        return varList;
    }




    private static String formatValue(ExcelEntity property, String cellValue) {
        if (property.getFormat() != null) {
            ExcelFormatBase<?> format = property.getFormat();
            try {
                if (format instanceof DefaultFormat) {
                    return ((DefaultFormat) format).ReadToExcel(String.class, cellValue).toString();
                }
                return format.ReadToExcel(cellValue, null).toString();
            } catch (Exception e) {
                throw new ExcelValueException(property.getTitle() + "错误", e);
            }
        }
        return cellValue;
    }

    private static Workbook getWorkbook(File file) throws IOException {
        FileInputStream fi = new FileInputStream(file);
        String fileType = file.getName().substring(file.getName().lastIndexOf(".") + 1);
        Workbook wb = null;
        if (fileType.equals("xls")) {
            wb = new HSSFWorkbook(fi);
        } else if (fileType.equals("xlsx")) {
            wb = new XSSFWorkbook(fi);
        } else {
            throw new IORunTimeException("您导入的文件不是标准excel文件");
        }
        return wb;
    }

    private static List<ExcelEntity> getExcelEntities(int startrow, int startcol, List<ExcelEntity> propertieList, Sheet sheet) {
        List<ExcelEntity> properties = new ArrayList<>();
        Row row = sheet.getRow(startrow - 1);
        int cellNum = row.getLastCellNum();
        Map<String, Integer> cellName = new HashMap<>();
        for (int j = startcol; j < cellNum; j++) {
            cellName.put(getCellValue(getMergedRegionValue(sheet, startrow - 1, j)), j);
        }
        for (ExcelEntity excelEntity : propertieList) {
            if ((cellName.containsKey(excelEntity.getTitle()) || (excelEntity.getAlias().isEmpty() && cellName.containsKey(excelEntity.getAlias()))) && excelEntity.isRead()) {
                excelEntity.setValue(cellName.get(excelEntity.getTitle()));
                properties.add(excelEntity);
            }
        }
        return properties;
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity,
                                        boolean safe, boolean readImage) {
        return readExcel(file, startrow, startcol, sheetnum, entity, safe, readImage, getExcelEntityList(entity));
    }

    public static <T> List<T> readExcel(File file, int startrow, int startcol, int sheetnum, Class<T> entity,
                                        boolean safe) {
        return readExcel(file, startrow, startcol, sheetnum, entity, safe, false, getExcelEntityList(entity));
    }

    public static <T> List<T> readExcel(String filepath, String filename, int startrow, int startcol, int sheetnum, Class<T> entity) {
        File target = new File(filepath, filename);
        return readExcel(target, startrow, startcol, sheetnum, entity);
    }

    public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity) {
        return readExcel(file.getFile(), startrow, startcol, sheetnum, entity);
    }

    public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity,
                                        boolean safe, boolean readImage) {
        return readExcel(file.getFile(), startrow, startcol, sheetnum, entity, safe, readImage);
    }

    public static <T> List<T> readExcel(UploadFile file, int startrow, int startcol, int sheetnum, Class<T> entity,
                                        boolean safe) {
        return readExcel(file.getFile(), startrow, startcol, sheetnum, entity, safe, false);
    }

    private static String getExcelStringValue(Workbook workbook, Cell cell, ExcelEntity property) {
        String cellValue = "";
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        if (null != cell) {
            cellValue = getCellValue(workbook, cell, cellValue, sdf, property.getAccuracy());
            if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
                throw new ExcelValueException(property.getTitle() + "为空");
            } else if (!cellValue.trim().isEmpty()) {
                if (property.getFormat() != null) {
                    ExcelFormatBase<?> format = property.getFormat();
                    try {
                        if (format instanceof DefaultFormat) {
                            return ((DefaultFormat) format).ReadToExcel(String.class, cellValue).toString();
                        }
                        return format.ReadToExcel(cellValue, null).toString();
                    } catch (Exception e) {
                        throw new ExcelValueException(property.getTitle() + "错误", e);
                    }
                }
                return cellValue;
            }
        }
        return null;
    }

    private static Object getCellValue(String v, ExcelEntity property, Map<String, String> objectMap) {
        Class<?> type = property.getType();
        ExcelFormatBase<?> format = property.getFormat();
        try {
            if (format instanceof DefaultFormat) {
                return ((DefaultFormat) format).ReadToExcel(type, v);
            }
            return format.ReadToExcel(v, objectMap);
        } catch (Exception e) {
            throw new ExcelValueException(property.getTitle() + "错误", e);
        }
    }

    private static String getCellStringValue(Workbook workbook, Cell cell, ExcelEntity property) {
        String cellValue = "";
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        if (null != cell) {
            cellValue = getCellValue(workbook, cell, cellValue, sdf, property.getAccuracy());
            if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
                throw new ExcelValueException(property.getTitle() + "为空");
            } else if (!cellValue.trim().isEmpty()) {
                return cellValue;
            }
        }
        return "";
    }

    private static Object getExcelValue(Workbook workbook, Cell cell, ExcelEntity property) {
        String cellValue = "";
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        if (null != cell) {
            cellValue = getCellValue(workbook, cell, cellValue, sdf, property.getAccuracy());
            if (property.isNotNull() && (cellValue == null || cellValue.trim().isEmpty())) {
                throw new ExcelValueException(property.getTitle() + "为空");
            } else if (!cellValue.trim().isEmpty()) {
                Class<?> type = property.getType();
                ExcelFormatBase<?> format = property.getFormat();
                try {
                    if (format instanceof DefaultFormat) {
                        return ((DefaultFormat) format).ReadToExcel(type, cellValue);
                    }
                    return format.ReadToExcel(cellValue, null);
                } catch (Exception e) {
                    throw new ExcelValueException(property.getTitle() + "错误", e);
                }
            } else {
                return null;
            }
        } else if (property.isNotNull()) {
            throw new ExcelValueException(property.getTitle() + "为空");
        }
        return null;
    }

    private static String getCellValue(Workbook workbook, Cell cell, String cellValue, SimpleDateFormat sdf, int accuracy) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = sdf.format(cell.getDateCellValue());
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue) && !Double.isInfinite(numericValue)) {
                        cellValue = String.valueOf((long) numericValue);
                    } else {
                        if (accuracy == -1) {
                            cellValue = new BigDecimal(String.valueOf(numericValue))
                                    .stripTrailingZeros()
                                    .toPlainString();
                        } else {
                            cellValue = BigDecimal.valueOf(numericValue)
                                    .setScale(accuracy, RoundingMode.HALF_UP)
                                    .toPlainString();
                        }
                    }
                }
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BLANK:
                cellValue = "";
                break;
            case BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                switch (evaluate.getCellType()) {
                    case NUMERIC:
                        cellValue = String.valueOf(evaluate.getNumberValue());
                        break;
                    case STRING:
                        cellValue = evaluate.getStringValue();
                        break;
                }
                break;
            case ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
        }
        return cellValue;
    }

    public static Cell getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    return fRow.getCell(firstColumn);
                }
            }
        }
        return sheet.getRow(row).getCell(column);
    }

    public static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        return cell.toString();
    }

    public static String getPictureExtension(Picture picture) {
        String mimeType = picture.getPictureData().getMimeType();
        return MIME_TYPE_TO_EXTENSION.getOrDefault(mimeType, "");
    }

    public static String checkNumber(String number) {
        String a = null;
        if (number.contains(".01") || number.contains(".02") || number.contains(".03") || number.contains(".04") || number.contains(".05")
                || number.contains(".06") || number.contains(".07") || number.contains(".08") || number.contains(".09")) {
            a = number;
        } else {
            if (number.contains(".0")) {
                a = number.substring(0, number.length() - 2);
            } else if (number.contains("-0")) {
                a = number;
            } else {
                a = number;
            }
        }
        return a;
    }

    public static <T> List<ExcelEntity> getExcelEntityList(Class<T> entity) {
        return ExcelBaseModel.getExcelEntities(entity);
    }
}
