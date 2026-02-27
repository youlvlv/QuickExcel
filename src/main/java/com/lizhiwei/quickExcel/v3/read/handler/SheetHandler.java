package com.lizhiwei.quickExcel.v3.read.handler;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.ParamType;
import com.lizhiwei.quickExcel.entity.ReadErrorInfo;
import com.lizhiwei.quickExcel.exception.ExcelReadException;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;
import com.lizhiwei.quickExcel.v3.read.converter.ValueConverter;
import com.lizhiwei.quickExcel.v3.read.model.CellData;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import com.lizhiwei.quickExcel.v3.read.parser.CellRefParser;
import com.lizhiwei.quickExcel.v3.read.parser.DispImgParser;
import com.lizhiwei.quickExcel.v3.read.parser.XmlImageParser;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * SAX 事件处理器 - 处理 Sheet 数据
 */
public class SheetHandler extends DefaultHandler {

    protected final SharedStrings sst;
    protected final Styles styles;
    protected final Class<?> entityClass;
    protected final List<ExcelEntity> properties;
    protected final int startRow;
    protected final int startCol;
    protected final List<Object> result;
    protected final boolean safe;
    protected final List<ReadErrorInfo> errors;

    protected String lastContents;
    protected boolean nextIsString;
    protected boolean inlineStr;
    protected int rowNum = 0;
    protected int colNum = 0;
    protected String cellRef;
    protected int formatIndex;
    protected String formatString;

    protected Object currentRow;
    protected Map<Integer, ExcelEntity> columnMapping;
    protected boolean headerParsed = false;
    protected final Map<Integer, String> headerRow = new HashMap<>();

    protected final List<CellData> currentRowCells;

    public SheetHandler(SharedStrings sst, Styles styles, Class<?> entityClass,
                        List<ExcelEntity> properties, int startRow, int startCol,
                        List result, boolean safe) {
        this.sst = sst;
        this.styles = styles;
        this.entityClass = entityClass;
        this.properties = properties;
        this.startRow = startRow - 1; // 转换为 0-based
        this.startCol = startCol;
        this.result = result;
        this.safe = safe;
        this.errors = new ArrayList<>();
        this.currentRowCells = new ArrayList<>();
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        if ("c".equals(name)) {
            cellRef = attributes.getValue("r");
            CellAddress cellAddress = new CellAddress(cellRef);
            rowNum = cellAddress.getRow();
            colNum = cellAddress.getColumn();

            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");

            if (cellType != null) {
                nextIsString = "s".equals(cellType);
                inlineStr = "inlineStr".equals(cellType);
            } else {
                nextIsString = false;
                inlineStr = false;
            }

            if (cellStyleStr != null && styles != null) {
                try {
                    int styleIndex = Integer.parseInt(cellStyleStr);
                    XSSFCellStyle style = styles.getStyleAt(styleIndex);
                    formatIndex = style.getDataFormat();
                    formatString = style.getDataFormatString();
                } catch (Exception e) {
                    formatIndex = 0;
                    formatString = "General";
                }
            } else {
                formatIndex = 0;
                formatString = "General";
            }

            lastContents = "";
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {
        if ("v".equals(name) || "t".equals(name)) {
            String value = lastContents;

            if (nextIsString && sst != null) {
                try {
                    int idx = Integer.parseInt(value);
                    value = sst.getItemAt(idx).toString();
                } catch (Exception e) {
                }
                nextIsString = false;
            }

            CellData cellData = createCellData(rowNum, colNum, value);
            processCellData(cellData);
            currentRowCells.add(cellData);
        }

        if ("row".equals(name)) {
            finishRow();
            currentRowCells.clear();
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        lastContents += new String(ch, start, length);
    }

    protected CellData createCellData(int row, int col, String value) {
        CellData cell = new CellData();
        cell.setRowIndex(row);
        cell.setColumnIndex(col);
        cell.setReference(CellRefParser.toReference(row, col));

        if (inlineStr) {
            cell.setCellType(CellData.CellType.INLINE_STRING);
            cell.setValue(value);
        } else if (nextIsString) {
            cell.setCellType(CellData.CellType.SHARED_STRING);
            cell.setRawValue(value);
        } else {
            cell.setCellType(CellData.CellType.NUMERIC);
            cell.setRawValue(value);
        }

        cell.setStyleIndex(formatIndex);
        cell.setFormatString(formatString);

        return cell;
    }

    protected void processCellData(CellData cellData) {
        int rowIndex = cellData.getRowIndex();
        int colIndex = cellData.getColumnIndex();
        String value = cellData.getRawValue();

        if (rowIndex == startRow) {
            headerRow.put(colIndex, value);
            if (!headerParsed && colIndex >= startCol) {
                if (columnMapping == null) {
                    columnMapping = new HashMap<>();
                }
                for (ExcelEntity prop : properties) {
                    if (prop.getTitle().equals(value) ||
                            (!prop.getAlias().isEmpty() && prop.getAlias().equals(value))) {
                        columnMapping.put(colIndex, prop);
                        break;
                    }
                }
            }
        }

        if (rowIndex > startRow && columnMapping != null) {
            if (!headerParsed) {
                headerParsed = true;
            }

            try {
                if (currentRow == null) {
                    currentRow = ValueConverter.newInstance(entityClass);
                }

                ExcelEntity property = columnMapping.get(colIndex);
                if (property != null) {
                    // 对于图片类型字段，调用专门的图片处理方法
                    if (property.getParamType() == ParamType.IMAGE) {
                        processCellWithImage(currentRow, rowIndex, colIndex, value, property);
                    } else {
                        ValueConverter.convertAndSet(currentRow, property, value);
                    }
                }
            } catch (Exception e) {
                handleError(rowIndex, "第" + (rowIndex + 1) + "行数据处理失败: " + e.getMessage(), e);
            }
        }
    }

    /**
     * 处理包含图片的单元格
     * 子类可以覆盖此方法以实现自定义的图片处理逻辑
     *
     * @param currentRow 当前行实体对象
     * @param rowNum     行号
     * @param colNum     列号
     * @param value      单元格值
     * @param property   实体属性
     */
    protected void processCellWithImage(Object currentRow, int rowNum, int colNum,
                                        String value, ExcelEntity property) {
        // 默认实现：尝试设置图片名称或值
        // 实际的图片数据需要在读取完成后通过位置匹配来设置
        try {
            // 如果是 DISPIMG 公式，提取图片名称
            if (value != null && value.contains("DISPIMG")) {
                // 暂时只设置公式值，图片数据后续处理
                DispImgParser.DispImgInfo info = DispImgParser.parse(value);
                XmlImageParser.CellImageRef cellRef = new XmlImageParser.CellImageRef(rowNum, colNum, null,
                        info.getImageName(), value);
                ImageData imageData =
                        XmlImageParser.parseDispImage(ExcelFileContextManager.getInstance().getThreadContext().get(),
                                cellRef);
                var saveFunction = Optional.ofNullable(property.getImageFileSaveFunction()).orElse(ExcelConfig.getImageFileFunction());
                ValueConverter.convertAndSet(currentRow, property, saveFunction.apply(imageData.getData(), imageData.getExtension()));
            }
        } catch (Exception e) {
            // 忽略错误
        }
    }

    protected void finishRow() {
        if (currentRow != null && rowNum >= startRow) {
            result.add(currentRow);
            currentRow = null;
        }
    }

    protected void handleError(int rowIndex, String message, Exception e) {
        if (safe) {
            errors.add(new ReadErrorInfo(rowIndex, message));
        } else {
            throw new ExcelReadException(message, e);
        }
    }

    public List<ReadErrorInfo> getErrors() {
        return errors;
    }

    /**
     * 通过反射设置字段值
     */
    protected void setFieldValue(Object obj, String fieldName, Object value) throws Exception {
        Field field = obj.getClass().getDeclaredField(fieldName);
        field.setAccessible(true);
        field.set(obj, value);
    }

    /**
     * 通过反射调用 Setter 方法
     */
    protected void setMethodValue(Object obj, String propertyName, Class<?> paramType, Object value) throws Exception {
        String setMethodName = "set" + propertyName.substring(0, 1).toUpperCase()
                + propertyName.substring(1);
        Method method = obj.getClass().getMethod(setMethodName, paramType);
        method.invoke(obj, value);
    }
}
