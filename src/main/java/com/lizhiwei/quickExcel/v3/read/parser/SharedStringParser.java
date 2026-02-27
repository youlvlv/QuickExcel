package com.lizhiwei.quickExcel.v3.read.parser;

import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipFile;
import java.util.zip.ZipEntry;
import java.io.InputStream;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

/**
 * 共享字符串解析器
 * 用于解析 xlsx 文件中的共享字符串表（xl/sharedStrings.xml）
 */
public class SharedStringParser {
    
    /**
     * 从 ZipFile 中读取共享字符串表
     */
    public static List<String> parse(ZipFile zipFile) {
        List<String> sharedStrings = new ArrayList<>();
        
        try {
            ZipEntry entry = zipFile.getEntry("xl/sharedStrings.xml");
            if (entry == null) {
                return sharedStrings;
            }
            
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            
            try (InputStream is = zipFile.getInputStream(entry)) {
                Document doc = builder.parse(is);
                NodeList siList = doc.getElementsByTagName("si");
                
                for (int i = 0; i < siList.getLength(); i++) {
                    Element si = (Element) siList.item(i);
                    String text = extractTextFromSi(si);
                    sharedStrings.add(text);
                }
            }
        } catch (Exception e) {
        }
        
        return sharedStrings;
    }
    
    /**
     * 从 <si> 元素中提取文本
     */
    public static String extractTextFromSi(Element si) {
        StringBuilder text = new StringBuilder();
        
        NodeList tList = si.getElementsByTagName("t");
        for (int i = 0; i < tList.getLength(); i++) {
            org.w3c.dom.Node t = tList.item(i);
            String parentName = t.getParentNode().getNodeName();
            if ("r".equals(parentName) || "si".equals(parentName)) {
                text.append(t.getTextContent());
            }
        }
        
        return text.toString();
    }
    
    /**
     * 获取共享字符串
     */
    public static String getString(List<String> sharedStrings, int index) {
        if (index >= 0 && index < sharedStrings.size()) {
            return sharedStrings.get(index);
        }
        return "";
    }
}
