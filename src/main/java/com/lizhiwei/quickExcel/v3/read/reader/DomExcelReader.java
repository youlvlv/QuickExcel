package com.lizhiwei.quickExcel.v3.read.reader;

import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ImageFileSaveFunction;
import com.lizhiwei.quickExcel.entity.ParamType;
import com.lizhiwei.quickExcel.entity.ReadErrorInfo;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;
import com.lizhiwei.quickExcel.v3.read.converter.ValueConverter;
import com.lizhiwei.quickExcel.v3.read.model.CellData;
import com.lizhiwei.quickExcel.v3.read.model.ExcelReadResult;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import com.lizhiwei.quickExcel.v3.read.parser.CellRefParser;
import com.lizhiwei.quickExcel.v3.read.parser.HeaderParser;
import com.lizhiwei.quickExcel.v3.read.parser.ImageParser;
import com.lizhiwei.quickExcel.v3.read.util.ImageSaveUtil;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * DOM 模式 Excel 读取器
 * <p>
 * 使用 {@link ExcelFileContext} 复用已解析的元数据，避免重复读取文件
 * </p>
 */
public class DomExcelReader extends AbstractExcelReader {

    private static final ThreadLocal<List<ReadErrorInfo>> readErrorInfoThreadLocal = ThreadLocal.withInitial(() -> new ArrayList<>());

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

            // 解析 sheet XML
            Document doc = parseSheet(sheetInfo.part);

            // 获取共享字符串（已缓存）
            List<String> sharedStrings = context.getSharedStringList();

            // 解析表头映射
            Map<Integer, ExcelEntity> columnMapping = parseHeaderMapping(doc, sharedStrings, properties, startRow, startCol);

            // 解析数据行
            result = parseDataRows(doc, sharedStrings, columnMapping, entity, startRow, startCol, safe);
            if (safe && !readErrorInfoThreadLocal.get().isEmpty()) {
                throw new ExcelReadException(readErrorInfoThreadLocal.get());
            }
        } catch (ExcelReadException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelReadException("读取 Excel 失败: " + e.getMessage(), e);
        } finally {
            readErrorInfoThreadLocal.remove();
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

            // 解析 sheet XML
            Document doc = parseSheet(sheetInfo.part);

            // 获取共享字符串（已缓存）
            List<String> sharedStrings = context.getSharedStringList();

            // 解析表头
            List<String> headerList = parseHeaderList(doc, sharedStrings, startRow, startCol);

            // 解析数据行
            result = parseDataRowsAsMap(doc, sharedStrings, headerList, startRow, startCol);

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

            // 解析 sheet XML
            Document doc = parseSheet(sheetInfo.part);

            // 获取共享字符串（已缓存）
            List<String> sharedStrings = context.getSharedStringList();

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
                            // 使用行和列的编码作为 key
                            int key = image.getRow() * 10000 + image.getColumn();

                            // 按属性名分组存储路径
                            columnImagePathMap
                                    .computeIfAbsent(prop.getProperty(), k -> new HashMap<>())
                                    .put(key, savedPath);
                        }
                    }
                }
            }

            // 解析表头映射
            Map<Integer, ExcelEntity> columnMapping = parseHeaderMapping(doc, sharedStrings, properties, startRow, startCol);

            // 解析数据行并绑定图片路径
            List<T> entities = parseDataRowsWithImages(doc, sharedStrings, columnMapping, entity,
                    startRow, startCol, safe, columnImagePathMap);

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
     * 解析 Sheet XML
     */
    private Document parseSheet(PackagePart sheetPart) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        DocumentBuilder builder = factory.newDocumentBuilder();

        try (InputStream is = sheetPart.getInputStream()) {
            return builder.parse(is);
        }
    }

    /**
     * 解析表头映射
     */
    private Map<Integer, ExcelEntity> parseHeaderMapping(Document doc, List<String> sharedStrings,
                                                         List<ExcelEntity> properties, int startRow, int startCol) {
        if (startRow <= 1) {
            return HeaderParser.createDefaultMapping(properties, startCol);
        }

        NodeList rowList = doc.getElementsByTagName("row");
        Element headerRow = findRowByIndex(rowList, startRow - 1);

        if (headerRow == null) {
            return HeaderParser.createDefaultMapping(properties, startCol);
        }

        List<CellData> headerCells = parseRowCells(headerRow, sharedStrings);
        return HeaderParser.parse(headerCells, properties, startCol);
    }

    /**
     * 解析表头列表
     */
    private List<String> parseHeaderList(Document doc, List<String> sharedStrings, int startRow, int startCol) {
        if (startRow <= 1) {
            return new ArrayList<>();
        }

        NodeList rowList = doc.getElementsByTagName("row");
        Element headerRow = findRowByIndex(rowList, startRow - 1);

        if (headerRow == null) {
            return new ArrayList<>();
        }

        List<CellData> headerCells = parseRowCells(headerRow, sharedStrings);
        return HeaderParser.parseAsList(headerCells, startCol);
    }

    /**
     * 根据行号查找 row 元素
     */
    private Element findRowByIndex(NodeList rowList, int rowIndex) {
        for (int i = 0; i < rowList.getLength(); i++) {
            Element row = (Element) rowList.item(i);
            String r = row.getAttribute("r");
            if (r != null && !r.isEmpty() && Integer.parseInt(r) == rowIndex) {
                return row;
            }
        }
        return null;
    }

    /**
     * 解析行单元格
     */
    private List<CellData> parseRowCells(Element row, List<String> sharedStrings) {
        List<CellData> cells = new ArrayList<>();
        NodeList cellList = row.getElementsByTagName("c");

        for (int i = 0; i < cellList.getLength(); i++) {
            Element cell = (Element) cellList.item(i);
            String ref = cell.getAttribute("r");
            int[] coords = CellRefParser.parse(ref);

            CellData cellData = new CellData();
            cellData.setRowIndex(coords[0]);
            cellData.setColumnIndex(coords[1]);
            cellData.setReference(ref);

            String type = cell.getAttribute("t");
            if ("s".equals(type)) {
                cellData.setCellType(CellData.CellType.SHARED_STRING);
            } else if ("inlineStr".equals(type)) {
                cellData.setCellType(CellData.CellType.INLINE_STRING);
            } else {
                cellData.setCellType(CellData.CellType.NUMERIC);
            }

            NodeList vList = cell.getElementsByTagName("v");
            if (vList.getLength() > 0) {
                String value = vList.item(0).getTextContent();
                if ("s".equals(type)) {
                    cellData.setRawValue(value);
                } else if ("inlineStr".equals(type)) {
                    NodeList isList = cell.getElementsByTagName("is");
                    if (isList.getLength() > 0) {
                        Element is = (Element) isList.item(0);
                        NodeList tList = is.getElementsByTagName("t");
                        if (tList.getLength() > 0) {
                            cellData.setValue(tList.item(0).getTextContent());
                        }
                    }
                } else {
                    cellData.setRawValue(value);
                }
            }

            cells.add(cellData);
        }

        return cells;
    }

    /**
     * 解析数据行
     */
    private <T> List<T> parseDataRows(Document doc, List<String> sharedStrings,
                                      Map<Integer, ExcelEntity> columnMapping,
                                      Class<T> entity, int startRow, int startCol, boolean safe) throws Exception {
        List<T> result = new ArrayList<>();
        NodeList rowList = doc.getElementsByTagName("row");

        for (int i = 0; i < rowList.getLength(); i++) {
            Element row = (Element) rowList.item(i);
            int rowIndex = getRowIndex(row);

            if (rowIndex < startRow) {
                continue;
            }

            T obj = ValueConverter.newInstance(entity);
            boolean hasData = false;

            List<CellData> cells = parseRowCells(row, sharedStrings);

            for (CellData cell : cells) {
                if (cell.getColumnIndex() < startCol) {
                    continue;
                }

                ExcelEntity property = columnMapping.get(cell.getColumnIndex());
                if (property != null) {
                    String value = resolveCellValue(cell, sharedStrings);
                    if (value != null && !value.isEmpty()) {
                        hasData = true;
                    }
                    try {
                        ValueConverter.convertAndSet(obj, property, value);
                    } catch (Exception e) {
                        if (safe) {
                            // 在 safe 模式下，收集错误但不中断
                            readErrorInfoThreadLocal.get().add(new ReadErrorInfo(rowIndex, "第" + (rowIndex + 1) +
                                    "行数据处理失败: " + e.getMessage()));
                        }
                    }
                }
            }

            if (hasData) {
                result.add(obj);
            }
        }

        return result;
    }

    /**
     * 解析数据行并绑定图片路径
     */
    private <T> List<T> parseDataRowsWithImages(Document doc, List<String> sharedStrings,
                                                Map<Integer, ExcelEntity> columnMapping,
                                                Class<T> entity, int startRow, int startCol,
                                                boolean safe, Map<String, Map<Integer, String>> columnImagePathMap) throws Exception {
        List<T> result = new ArrayList<>();
        NodeList rowList = doc.getElementsByTagName("row");

        for (int i = 0; i < rowList.getLength(); i++) {
            Element row = (Element) rowList.item(i);
            int rowIndex = getRowIndex(row);

            if (rowIndex < startRow) {
                continue;
            }

            T obj = ValueConverter.newInstance(entity);
            boolean hasData = false;

            List<CellData> cells = parseRowCells(row, sharedStrings);

            for (CellData cell : cells) {
                if (cell.getColumnIndex() < startCol) {
                    continue;
                }

                ExcelEntity property = columnMapping.get(cell.getColumnIndex());
                if (property != null) {
                    // 如果是图片类型字段，查找对应的图片路径
                    if (property.getParamType() == ParamType.IMAGE) {
                        Map<Integer, String> pathMap = columnImagePathMap.get(property.getProperty());

                        if (pathMap != null) {
                            int key = rowIndex * 10000 + cell.getColumnIndex();
                            String imagePath = pathMap.get(key);

                            if (imagePath != null && !imagePath.isEmpty()) {
                                hasData = true;
                                try {
                                    setFieldValue(obj, property, imagePath);
                                    continue;
                                } catch (Exception e) {
                                    // 忽略错误
                                }
                            }
                        }
                    } else {
                        String value = resolveCellValue(cell, sharedStrings);
                        if (value != null && !value.isEmpty()) {
                            hasData = true;
                        }
                        try {
                            ValueConverter.convertAndSet(obj, property, value);
                        } catch (Exception e) {
                            if (safe) {
                                // 在 safe 模式下，收集错误但不中断
                            }
                        }
                    }
                }
            }

            if (hasData) {
                result.add(obj);
            }
        }

        return result;
    }

    /**
     * 解析数据行为 Map
     */
    private List<Map<String, String>> parseDataRowsAsMap(Document doc, List<String> sharedStrings,
                                                         List<String> headers, int startRow, int startCol) {
        List<Map<String, String>> result = new ArrayList<>();
        NodeList rowList = doc.getElementsByTagName("row");

        for (int i = 0; i < rowList.getLength(); i++) {
            Element row = (Element) rowList.item(i);
            int rowIndex = getRowIndex(row);

            if (rowIndex < startRow) {
                continue;
            }

            Map<String, String> rowData = new HashMap<>();
            boolean hasData = false;

            List<CellData> cells = parseRowCells(row, sharedStrings);

            for (CellData cell : cells) {
                if (cell.getColumnIndex() < startCol) {
                    continue;
                }

                String value = resolveCellValue(cell, sharedStrings);
                if (value != null && !value.isEmpty()) {
                    hasData = true;
                }

                String header = getHeader(headers, cell.getColumnIndex(), startCol);
                rowData.put(header, value);
            }

            if (hasData) {
                result.add(rowData);
            }
        }

        return result;
    }

    /**
     * 解析单元格值
     */
    private String resolveCellValue(CellData cell, List<String> sharedStrings) {
        CellData.CellType cellType = cell.getCellType();

        if (cellType == CellData.CellType.SHARED_STRING) {
            String rawValue = cell.getRawValue();
            if (rawValue != null && !rawValue.isEmpty()) {
                try {
                    int idx = Integer.parseInt(rawValue);
                    if (idx >= 0 && idx < sharedStrings.size()) {
                        return sharedStrings.get(idx);
                    }
                    return rawValue;
                } catch (NumberFormatException e) {
                    return rawValue;
                }
            }
        } else if (cellType == CellData.CellType.INLINE_STRING) {
            return cell.getValue() != null ? cell.getValue() : "";
        }

        return cell.getRawValue() != null ? cell.getRawValue() : "";
    }

    /**
     * 获取表头
     */
    private String getHeader(List<String> headers, int colIndex, int startCol) {
        int headerIndex = colIndex - startCol;
        if (headerIndex >= 0 && headerIndex < headers.size()) {
            String header = headers.get(headerIndex);
            if (header != null && !header.isEmpty()) {
                return header;
            }
        }
        return "Column" + (colIndex + 1);
    }

    /**
     * 获取行索引
     */
    private int getRowIndex(Element row) {
        String r = row.getAttribute("r");
        if (r != null && !r.isEmpty()) {
            return Integer.parseInt(r);
        }
        return 0;
    }

    /**
     * 通过反射设置字段值
     */
    private void setFieldValue(Object obj, ExcelEntity property, Object value) throws Exception {
        java.lang.reflect.Field field = obj.getClass().getDeclaredField(property.getProperty());
        field.setAccessible(true);
        field.set(obj, value);
    }
}
