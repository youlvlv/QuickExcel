package com.lizhiwei.quickExcel.core;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.v2.reader.PoiExcelReader;
import com.lizhiwei.quickExcel.v3.reader.XmlExcelReaderAdapter;

/**
 * Excel 读取器工厂
 * 根据配置创建对应的读取器实例
 */
public class ExcelReaderFactory {
    
    private static volatile ExcelReader v2Reader;
    private static volatile ExcelReader v3Reader;
    
    /**
     * 获取默认读取器
     * 根据 ExcelConfig 中的配置返回 V2 或 V3 读取器
     * @return ExcelReader 实例
     */
    public static ExcelReader getReader() {
        return getReader(ExcelConfig.getDefaultReadEngine());
    }
    
    /**
     * 获取指定版本的读取器
     * @param engine 引擎版本
     * @return ExcelReader 实例
     */
    public static ExcelReader getReader(ExcelConfig.ReadEngine engine) {
        return switch (engine) {
            case V2 -> getV2Reader();
            case V3 -> getV3Reader();
            default -> getV2Reader();
        };
    }
    
    /**
     * 获取 V2 读取器（Apache POI）
     */
    public static ExcelReader getV2Reader() {
        if (v2Reader == null) {
            synchronized (ExcelReaderFactory.class) {
                if (v2Reader == null) {
                    v2Reader = new PoiExcelReader();
                }
            }
        }
        return v2Reader;
    }
    
    /**
     * 获取 V3 读取器（SAX/DOM）
     */
    public static ExcelReader getV3Reader() {
        if (v3Reader == null) {
            synchronized (ExcelReaderFactory.class) {
                if (v3Reader == null) {
                    v3Reader = new XmlExcelReaderAdapter();
                }
            }
        }
        return v3Reader;
    }
    
    /**
     * 设置自定义的 V2 读取器（用于扩展或测试）
     */
    public static void setV2Reader(ExcelReader reader) {
        v2Reader = reader;
    }
    
    /**
     * 设置自定义的 V3 读取器（用于扩展或测试）
     */
    public static void setV3Reader(ExcelReader reader) {
        v3Reader = reader;
    }
}
