package com.lizhiwei.quickExcel.v3.read.util;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Workbook XML 解析器
 * 用于解析 xl/workbook.xml 获取 sheet 名称和索引的映射
 */
public class WorkbookXmlParser {
    
    /**
     * Sheet 信息
     */
    public static class SheetInfo {
        public final int index;
        public final String name;
        public final String relId;
        public final int sheetId;
        
        public SheetInfo(int index, String name, String relId, int sheetId) {
            this.index = index;
            this.name = name;
            this.relId = relId;
            this.sheetId = sheetId;
        }
    }
    
    /**
     * 解析 workbook.xml 获取所有 sheet 信息
     * @param pkg OPCPackage
     * @return sheet 名称到信息的映射
     */
    public static Map<String, SheetInfo> parseSheetInfos(OPCPackage pkg) {
        Map<String, SheetInfo> sheetMap = new HashMap<>();
        
        try {
            PackagePart workbookPart = getWorkbookPart(pkg);
            if (workbookPart == null) {
                return sheetMap;
            }
            
            try (InputStream is = workbookPart.getInputStream()) {
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(is);
                
                NodeList sheetNodes = doc.getElementsByTagNameNS("*", "sheet");
                for (int i = 0; i < sheetNodes.getLength(); i++) {
                    Element sheetElement = (Element) sheetNodes.item(i);
                    
                    String name = sheetElement.getAttribute("name");
                    String relId = sheetElement.getAttributeNS(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                    if (relId.isEmpty()) {
                        relId = sheetElement.getAttribute("r:id");
                    }
                    
                    String sheetIdStr = sheetElement.getAttribute("sheetId");
                    int sheetId = sheetIdStr.isEmpty() ? i + 1 : Integer.parseInt(sheetIdStr);
                    
                    SheetInfo info = new SheetInfo(i, name, relId, sheetId);
                    sheetMap.put(name, info);
                }
            }
        } catch (Exception e) {
            // 解析失败返回空映射
        }
        
        return sheetMap;
    }
    
    /**
     * 根据 sheet 名称获取索引
     * @param pkg OPCPackage
     * @param sheetName sheet 名称
     * @return sheet 索引，找不到返回 -1
     */
    public static int getSheetIndex(OPCPackage pkg, String sheetName) {
        Map<String, SheetInfo> sheetMap = parseSheetInfos(pkg);
        SheetInfo info = sheetMap.get(sheetName);
        return info != null ? info.index : -1;
    }
    
    /**
     * 根据 sheet 索引获取名称
     * @param pkg OPCPackage
     * @param sheetIndex sheet 索引
     * @return sheet 名称，找不到返回 null
     */
    public static String getSheetName(OPCPackage pkg, int sheetIndex) {
        try {
            PackagePart workbookPart = getWorkbookPart(pkg);
            if (workbookPart == null) {
                return null;
            }
            
            try (InputStream is = workbookPart.getInputStream()) {
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(is);
                
                NodeList sheetNodes = doc.getElementsByTagNameNS("*", "sheet");
                if (sheetIndex >= 0 && sheetIndex < sheetNodes.getLength()) {
                    Element sheetElement = (Element) sheetNodes.item(sheetIndex);
                    return sheetElement.getAttribute("name");
                }
            }
        } catch (Exception e) {
            // 解析失败
        }
        return null;
    }
    
    /**
     * 获取所有 sheet 名称列表
     * @param pkg OPCPackage
     * @return sheet 名称列表
     */
    public static List<String> getSheetNames(OPCPackage pkg) {
        List<String> names = new ArrayList<>();
        try {
            PackagePart workbookPart = getWorkbookPart(pkg);
            if (workbookPart == null) {
                return names;
            }
            
            try (InputStream is = workbookPart.getInputStream()) {
                DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                factory.setNamespaceAware(true);
                DocumentBuilder builder = factory.newDocumentBuilder();
                Document doc = builder.parse(is);
                
                NodeList sheetNodes = doc.getElementsByTagNameNS("*", "sheet");
                for (int i = 0; i < sheetNodes.getLength(); i++) {
                    Element sheetElement = (Element) sheetNodes.item(i);
                    names.add(sheetElement.getAttribute("name"));
                }
            }
        } catch (Exception e) {
            // 解析失败
        }
        return names;
    }
    
    /**
     * 获取 workbook.xml 的 PackagePart
     */
    private static PackagePart getWorkbookPart(OPCPackage pkg) {
        try {
            for (PackagePart part : pkg.getParts()) {
                if ("/xl/workbook.xml".equals(part.getPartName().getName())) {
                    return part;
                }
            }
        } catch (Exception e) {
            // 忽略错误
        }
        return null;
    }
}
