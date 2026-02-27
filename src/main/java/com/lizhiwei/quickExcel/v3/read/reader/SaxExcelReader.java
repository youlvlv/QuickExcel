package com.lizhiwei.quickExcel.v3.read.reader;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ImageFileSaveFunction;
import com.lizhiwei.quickExcel.entity.ParamType;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;
import com.lizhiwei.quickExcel.v3.read.handler.SheetHandler;
import com.lizhiwei.quickExcel.v3.read.model.ExcelReadResult;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import com.lizhiwei.quickExcel.v3.read.parser.ImageParser;
import com.lizhiwei.quickExcel.v3.read.util.ImageSaveUtil;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.xml.sax.InputSource;
import org.xml.sax.Attributes;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.BiFunction;

/**
 * SAX 模式 Excel 读取器
 * <p>
 * 使用 {@link ExcelFileContext} 复用已解析的元数据，避免重复读取文件
 * </p>
 */
public class SaxExcelReader extends AbstractExcelReader {
    
    @Override
    protected <T> List<T> doReadExcel(File file, int startRow, int startCol, int sheetNum,
                                       Class<T> entity, List<ExcelEntity> properties, boolean safe) {
        // 获取缓存的文件上下文
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        
        List<T> result = new ArrayList<>();
        
        try {
            // 获取指定 sheet
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetNum);
            if (sheetInfo == null) {
                throw new ExcelReadException("Sheet 索引超出范围: " + sheetNum);
            }
            
            // 获取共享字符串表（已缓存）
            SharedStrings sst = context.getSharedStrings();
            
            // 尝试获取样式表（可能为 null）
            Styles styles = null;
            try {
                styles = context.getStyles();
            } catch (Exception e) {
                // 样式非必需
            }
            
            // 使用 sheet 的 InputStream 进行 SAX 解析
            try (InputStream sheetStream = sheetInfo.part.getInputStream()) {
                XMLReader parser = XMLReaderFactory.createXMLReader();
                SheetHandler handler = new SheetHandler(sst, styles, entity, properties, startRow, startCol, result, safe);
                parser.setContentHandler(handler);
                
                parser.parse(new InputSource(sheetStream));
                
                if (safe && !handler.getErrors().isEmpty()) {
                    throw new ExcelReadException(handler.getErrors());
                }
            }
            
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("读取 Excel 失败: " + e.getMessage(), e);
        }
        
        return result;
    }
    
    @Override
    protected List<Map<String, String>> doReadExcelAsMap(File file, int startRow, int startCol,
                                                         int sheetNum, List<String> headers) {
        // 获取缓存的文件上下文
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        
        List<Map<String, String>> result = new ArrayList<>();
        
        try {
            // 获取指定 sheet
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetNum);
            if (sheetInfo == null) {
                throw new ExcelReadException("Sheet 索引超出范围: " + sheetNum);
            }
            
            // 获取共享字符串表（已缓存）
            SharedStrings sst = context.getSharedStrings();
            
            Styles styles = null;
            try {
                styles = context.getStyles();
            } catch (Exception e) {
                // 样式非必需
            }
            
            // 使用 sheet 的 InputStream 进行 SAX 解析
            try (InputStream sheetStream = sheetInfo.part.getInputStream()) {
                XMLReader parser = XMLReaderFactory.createXMLReader();
                MapSheetHandler handler = new MapSheetHandler(sst, styles, headers, startRow, startCol, result);
                parser.setContentHandler(handler);
                
                parser.parse(new InputSource(sheetStream));
            }
            
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("读取 Excel 失败: " + e.getMessage(), e);
        }
        
        return result;
    }
    
    @Override
    public List<ImageData> readExcelImages(File file, int sheetNum) {
        // 使用上下文管理器获取缓存的上下文
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        return ImageParser.parse(context, sheetNum);
    }
    
    @Override
    protected <T> ExcelReadResult<T> doReadExcelWithImages(File file, int startRow, int startCol, 
                                                            int sheetNum, Class<T> entity, 
                                                            List<ExcelEntity> properties, boolean safe) {
        ExcelReadResult<T> result = new ExcelReadResult<>();
        
        // 获取缓存的文件上下文
        ExcelFileContext context = ExcelFileContextManager.getInstance().getContext(file);
        
        try {
            // 获取指定 sheet
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetNum);
            if (sheetInfo == null) {
                throw new ExcelReadException("Sheet 索引超出范围: " + sheetNum);
            }
            
            // 获取共享字符串表（已缓存）
            SharedStrings sst = context.getSharedStrings();
            
            // 尝试获取样式表（可能为 null）
            Styles styles = null;
            try {
                styles = context.getStyles();
            } catch (Exception e) {
                // 样式非必需
            }
            
            // 【关键】获取所有图片（浮动图片 + DISPIMG 图片）
            List<ImageData> images = ImageParser.parse(context, sheetNum);
            
            // 保存图片并构建位置到路径的映射
            // 每个图片字段可能有不同的保存函数
            Map<String, Map<Integer, String>> columnImagePathMap = new HashMap<>();
            
            for (ImageData image : images) {
                // 添加到结果中
                result.addImage(image);
                
                // 为该图片在每个图片字段中查找对应的保存函数
                for (ExcelEntity prop : properties) {
                    if (prop.getParamType() == ParamType.IMAGE) {
                        // 获取该字段的保存函数（优先使用字段配置，否则使用全局配置）
                        ImageFileSaveFunction saveFunction = prop.getImageFileSaveFunction();
                        
                        // 保存图片
                        String savedPath = ImageSaveUtil.saveImage(image, saveFunction);
                        
                        if (savedPath != null && !savedPath.isEmpty()) {
                            String key = image.getRow() + "," + image.getColumn();
                            
                            // 按属性名分组存储路径
                            columnImagePathMap
                                .computeIfAbsent(prop.getProperty(), k -> new HashMap<>())
                                .put(image.getRow() * 10000 + image.getColumn(), savedPath);
                        }
                    }
                }
            }
            
            // 使用 sheet 的 InputStream 进行 SAX 解析
            List<T> entities = new ArrayList<>();
            try (InputStream sheetStream = sheetInfo.part.getInputStream()) {
                XMLReader parser = XMLReaderFactory.createXMLReader();
                ImageAwareSheetHandler<T> handler = new ImageAwareSheetHandler<>(
                        sst, styles, entity, properties, startRow, startCol, 
                        entities, safe, columnImagePathMap);
                parser.setContentHandler(handler);
                
                parser.parse(new InputSource(sheetStream));
                
                if (safe && !handler.getErrors().isEmpty()) {
                    throw new ExcelReadException(handler.getErrors());
                }
            }
            
            // 将解析的实体添加到结果
            for (T entityObj : entities) {
                result.addEntity(entityObj);
            }
            
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("读取 Excel 失败: " + e.getMessage(), e);
        }
        
        return result;
    }
    
    /**
     * 支持图片保存的 SheetHandler
     */
    private static class ImageAwareSheetHandler<T> extends SheetHandler {
        private final Map<String, Map<Integer, String>> columnImagePathMap;
        private final List<ExcelEntity> imageProperties;
        
        public ImageAwareSheetHandler(SharedStrings sst, Styles styles, Class<T> entityClass, 
                                       List<ExcelEntity> properties, int startRow, int startCol, 
                                       List<T> resultList, boolean safe,
                                       Map<String, Map<Integer, String>> columnImagePathMap) {
            super(sst, styles, entityClass, properties, startRow, startCol, resultList, safe);
            this.columnImagePathMap = columnImagePathMap;
            
            // 筛选出图片类型的属性
            this.imageProperties = new ArrayList<>();
            for (ExcelEntity prop : properties) {
                if (prop.getParamType() == ParamType.IMAGE) {
                    imageProperties.add(prop);
                }
            }
        }
        
        @Override
        protected void processCellWithImage(Object currentRow, int rowNum, int colNum, 
                                             String value, ExcelEntity property) {
            if (property != null && property.getParamType() == ParamType.IMAGE) {
                // 查找该单元格位置、该属性的图片路径
                Map<Integer, String> pathMap = columnImagePathMap.get(property.getProperty());
                
                if (pathMap != null) {
                    // 使用行和列的编码作为 key
                    int key = rowNum * 10000 + colNum;
                    String imagePath = pathMap.get(key);
                    
                    if (imagePath != null && !imagePath.isEmpty()) {
                        try {
                            // 将图片路径设置到实体字段
                            setFieldValue(currentRow, property, imagePath);
                            return;
                        } catch (Exception e) {
                            // 忽略错误
                        }
                    }
                }
                
                // 没有图片或保存失败，尝试设置原始值
                try {
                    super.processCellWithImage(currentRow, rowNum, colNum, value, property);
                } catch (Exception ex) {
                    // 忽略错误
                }
            } else {
                super.processCellWithImage(currentRow, rowNum, colNum, value, property);
            }
        }
        
        private void setFieldValue(Object obj, ExcelEntity property, Object value) throws Exception {
            switch (property.getParamType()) {
                case FIELD:
                case IMAGE:
                    java.lang.reflect.Field field = obj.getClass().getDeclaredField(property.getProperty());
                    field.setAccessible(true);
                    field.set(obj, value);
                    break;
                case METHOD:
                    String setMethodName = "set" + property.getProperty().substring(0, 1).toUpperCase() 
                            + property.getProperty().substring(1);
                    java.lang.reflect.Method method = obj.getClass().getMethod(setMethodName, property.getType());
                    method.invoke(obj, value);
                    break;
            }
        }
    }
    
    /**
     * Map 解析的 SheetHandler
     */
    private static class MapSheetHandler extends SheetHandler {
        private final List<String> headers;
        private final int startRow;
        private final int startCol;
        private final List<Map<String, String>> result;
        
        private String lastContents;
        private boolean nextIsString;
        private boolean inlineStr;
        private int rowNum = 0;
        private int colNum = 0;
        
        private Map<String, String> currentRow;
        private List<String> actualHeaders = new ArrayList<>();
        
        private final SharedStrings sharedStrings;
        
        public MapSheetHandler(SharedStrings sst, Styles styles, List<String> headers,
                               int startRow, int startCol, List<Map<String, String>> result) {
            super(sst, styles, Object.class, new ArrayList<>(), startRow, startCol, new ArrayList<>(), false);
            this.sharedStrings = sst;
            this.headers = headers;
            this.startRow = startRow - 1;
            this.startCol = startCol;
            this.result = result;
        }
        
        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if ("c".equals(name)) {
                String cellRef = attributes.getValue("r");
                org.apache.poi.ss.util.CellAddress cellAddress = new org.apache.poi.ss.util.CellAddress(cellRef);
                rowNum = cellAddress.getRow();
                colNum = cellAddress.getColumn();
                
                String cellType = attributes.getValue("t");
                if (cellType != null) {
                    nextIsString = "s".equals(cellType);
                    inlineStr = "inlineStr".equals(cellType);
                } else {
                    nextIsString = false;
                    inlineStr = false;
                }
                
                lastContents = "";
            }
        }
        
        @Override
        public void endElement(String uri, String localName, String name) {
            if ("v".equals(name) || "t".equals(name)) {
                String value = lastContents;
                
                if (nextIsString && sharedStrings != null) {
                    try {
                        int idx = Integer.parseInt(value);
                        value = sharedStrings.getItemAt(idx).toString();
                    } catch (Exception ignored) {
                    }
                    nextIsString = false;
                }
                
                if (rowNum == startRow && colNum >= startCol) {
                    while (actualHeaders.size() <= colNum - startCol) {
                        actualHeaders.add("");
                    }
                    actualHeaders.set(colNum - startCol, value);
                }
                
                if (rowNum >= startRow && colNum >= startCol) {
                    if (currentRow == null) {
                        currentRow = new java.util.LinkedHashMap<>();
                    }
                    
                    int headerIndex = colNum - startCol;
                    String header = headerIndex < actualHeaders.size() ?
                            actualHeaders.get(headerIndex) : "Column" + (colNum + 1);
                    
                    if (headers == null || headers.isEmpty() || headers.contains(header)) {
                        currentRow.put(header, value);
                    }
                }
            }
            
            if ("row".equals(name)) {
                if (currentRow != null && rowNum >= startRow && !currentRow.isEmpty()) {
                    result.add(currentRow);
                }
                currentRow = null;
            }
        }
        
        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }
    }
}
