package com.lizhiwei.quickExcel.v3.read.parser;

import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 基于 XML 解析的图片解析器
 * 不依赖 XSSFWorkbook，直接解析 xlsx 内部的 XML 结构
 * <p>
 * 支持两种模式：
 * 1. 直接传入 File 解析（一次性）
 * 2. 传入 ExcelFileContext 复用已解析的元数据
 * 3. 支持 DISPIMG 函数图片（WPS 嵌入单元格图片）
 * </p>
 */
public class XmlImageParser {
    
    private static final Map<String, String> MIME_TYPE_TO_EXTENSION = new HashMap<>();
    
    static {
        MIME_TYPE_TO_EXTENSION.put("image/png", "png");
        MIME_TYPE_TO_EXTENSION.put("image/jpeg", "jpg");
        MIME_TYPE_TO_EXTENSION.put("image/gif", "gif");
        MIME_TYPE_TO_EXTENSION.put("image/bmp", "bmp");
        MIME_TYPE_TO_EXTENSION.put("image/x-emf", "emf");
        MIME_TYPE_TO_EXTENSION.put("image/x-wmf", "wmf");
    }
    
    // ==================== 使用上下文解析（推荐）====================
    
    /**
     * 使用上下文解析图片
     * @param context 文件上下文（已包含解析好的元数据）
     * @param sheetNum sheet 索引
     * @return 图片列表（包含浮动图片和 DISPIMG 图片）
     */
    public static List<ImageData> parse(ExcelFileContext context, int sheetNum) {
        List<ImageData> images = new ArrayList<>();
        try {
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetNum);
            if (sheetInfo == null) {
                return images;
            }
            
            // 1. 解析浮动图片（传统方式）
            images.addAll(parseFloatingImages(context, sheetInfo));
            
            // 2. 解析 DISPIMG 图片（WPS 嵌入单元格图片）
            images.addAll(parseDispImages(context, sheetInfo));
            
        } catch (Exception e) {
            // 忽略错误
        }
        return images;
    }
    
    /**
     * 使用上下文解析图片（按 sheet 名称）
     * @param context 文件上下文
     * @param sheetName sheet 名称
     * @return 图片列表
     */
    public static List<ImageData> parse(ExcelFileContext context, String sheetName) {
        try {
            ExcelFileContext.SheetInfo sheetInfo = context.getSheet(sheetName);
            if (sheetInfo == null) {
                return new ArrayList<>();
            }
            return parse(context, sheetInfo.index);
        } catch (Exception e) {
            return new ArrayList<>();
        }
    }
    
    /**
     * 解析指定 sheet 的浮动图片
     */
    private static List<ImageData> parseFloatingImages(ExcelFileContext context, ExcelFileContext.SheetInfo sheetInfo) {
        List<ImageData> images = new ArrayList<>();
        
        try {
            // 查找 drawing 关系
            for (PackageRelationship rel : sheetInfo.part.getRelationships()) {
                if (rel.getRelationshipType().contains("drawing")) {
                    PackagePart drawingPart = context.getPackage().getPart(rel);
                    if (drawingPart != null) {
                        List<ImageRef> refs = parseDrawingXml(drawingPart);
                        for (ImageRef ref : refs) {
                            ImageData imageData = createImageData(context, ref);
                            if (imageData != null) {
                                images.add(imageData);
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            // 忽略错误
        }
        
        return images;
    }
    
    /**
     * 解析 DISPIMG 图片（WPS 嵌入单元格图片）
     * <p>
     * 正确的解析流程：
     * 1. 从单元格公式获取图片名称（如 _xlsInner_1 或 ID_D02D9A6B933B477E93E74492C3957BA7）
     * 2. 从 xl/cellimages.xml 根据图片名称找对应的 cellImage 元素，获取 a:blip 的 r:embed 属性（rId）
     * 3. 从 xl/_rels/cellimages.xml.rels 根据 rId 找 Target 路径
     * 4. 根据 Target 路径获取实际图片文件
     * </p>
     */
    private static List<ImageData> parseDispImages(ExcelFileContext context, ExcelFileContext.SheetInfo sheetInfo) {
        List<ImageData> images = new ArrayList<>();
        
        try {
            // 解析 sheet XML，查找包含 DISPIMG 公式的单元格
            Map<String, CellImageRef> dispImgRefs = parseSheetForDispImg(sheetInfo);
            
            if (dispImgRefs.isEmpty()) {
                return images;
            }
            
            // 获取缓存的 DISPIMG 元数据
            Map<String, String> imageNameToRid = context.getCellImagesNameToRid();
            Map<String, String> ridToTarget = context.getCellImagesRidToTarget();
            
            // 根据元数据获取图片
            for (CellImageRef ref : dispImgRefs.values()) {
                ImageData image = findDispImage(context, ref, imageNameToRid, ridToTarget);
                if (image != null) {
                    images.add(image);
                }
            }
            
        } catch (Exception e) {
            // 忽略错误
        }
        
        return images;
    }

    public static ImageData parseDispImage(ExcelFileContext context, CellImageRef ref) {
        try {
            // 使用缓存的 DISPIMG 元数据获取图片
            Map<String, String> imageNameToRid = context.getCellImagesNameToRid();
            Map<String, String> ridToTarget = context.getCellImagesRidToTarget();
            return findDispImage(context, ref, imageNameToRid, ridToTarget);
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 解析 sheet XML，查找 DISPIMG 公式
     */
    private static Map<String, CellImageRef> parseSheetForDispImg(ExcelFileContext.SheetInfo sheetInfo) {
        Map<String, CellImageRef> refs = new HashMap<>();
        
        try (InputStream is = sheetInfo.part.getInputStream()) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(is);
            
            // 查找所有包含公式的单元格
            NodeList cells = doc.getElementsByTagNameNS("*", "c");
            for (int i = 0; i < cells.getLength(); i++) {
                Element cell = (Element) cells.item(i);
                
                // 获取单元格引用
                String ref = cell.getAttribute("r");
                if (ref.isEmpty()) continue;
                
                // 解析行号和列号
                int[] coords = parseCellRef(ref);
                
                // 查找公式节点
                NodeList formulas = cell.getElementsByTagNameNS("*", "f");
                if (formulas.getLength() > 0) {
                    String formula = formulas.item(0).getTextContent();
                    
                    // 检查是否为 DISPIMG 公式
                    if (DispImgParser.isDispImgFormula(formula)) {
                        DispImgParser.DispImgInfo info = DispImgParser.parse(formula);
                        if (info != null) {
                            CellImageRef cellRef = new CellImageRef(coords[0], coords[1], ref, info.getImageName(), formula);
                            refs.put(ref, cellRef);
                        }
                    }
                }
            }
            
        } catch (Exception e) {
            // 解析失败
        }
        
        return refs;
    }
    
    /**
     * 查找 DISPIMG 对应的图片
     * <p>
     * 正确流程：
     * 1. 从 imageName 通过 imageNameToRid 获取 rId
     * 2. 从 rId 通过 ridToTarget 获取 Target 路径
     * 3. 根据 Target 路径获取图片 Part
     * </p>
     */
    private static ImageData findDispImage(ExcelFileContext context, CellImageRef ref, 
                                           Map<String, String> imageNameToRid,
                                           Map<String, String> ridToTarget) {
        try {
            String imageName = ref.imageName;
            
            // 1. 从图片名称获取 rId
            String rid = imageNameToRid.get(imageName);
            if (rid == null) {
                // 尝试查找可能的匹配（图片名称可能被修改过）
                for (Map.Entry<String, String> entry : imageNameToRid.entrySet()) {
                    if (entry.getKey().contains(imageName) || imageName.contains(entry.getKey())) {
                        rid = entry.getValue();
                        break;
                    }
                }
            }
            
            if (rid == null) {
                return null;
            }
            
            // 2. 从 rId 获取 Target 路径
            String target = ridToTarget.get(rid);
            if (target == null) {
                return null;
            }
            
            // 3. 根据 Target 路径获取图片 Part
            // Target 路径可能是相对路径（如 "media/image1.png"），需要转换为绝对路径
            String absolutePath = target;
            if (!target.startsWith("/")) {
                // 相对于 xl 目录
                absolutePath = "/xl/" + target;
            }
            
            PackagePartName partName = PackagingURIHelper.createPartName(absolutePath);
            PackagePart imagePart = context.getPackage().getPart(partName);
            
            if (imagePart == null) {
                // 尝试从所有 parts 中查找
                for (PackagePart part : context.getPackage().getParts()) {
                    String name = part.getPartName().getName();
                    if (name.endsWith(target) || name.contains(target)) {
                        imagePart = part;
                        break;
                    }
                }
            }
            
            if (imagePart != null) {
                // 获取扩展名
                String contentType = imagePart.getContentType();
                String extension = MIME_TYPE_TO_EXTENSION.getOrDefault(contentType, "png");
                
                // 创建 DISPIMG 类型的图片数据
                ImageData imageData = ImageData.createDispImg(ref.row, ref.col, imageName, imagePart.getInputStream(), extension);
                imageData.setFormula(ref.formula);
                imageData.setPath(imagePart.getPartName().getName());
                
                return imageData;
            }
            
        } catch (Exception e) {
            // 查找失败
        }
        
        return null;
    }
    
    /**
     * 解析单元格引用（如 "A1" -> [0, 0]）
     */
    private static int[] parseCellRef(String ref) {
        int row = 0;
        int col = 0;
        
        int i = 0;
        // 解析列（字母部分）
        while (i < ref.length() && Character.isLetter(ref.charAt(i))) {
            col = col * 26 + (Character.toUpperCase(ref.charAt(i)) - 'A' + 1);
            i++;
        }
        col--; // 转换为 0-based
        
        // 解析行（数字部分）
        while (i < ref.length() && Character.isDigit(ref.charAt(i))) {
            row = row * 10 + (ref.charAt(i) - '0');
            i++;
        }
        row--; // 转换为 0-based
        
        return new int[]{row, col};
    }
    
    // ==================== 直接文件解析（兼容旧方式）====================
    
    /**
     * 从 xlsx 文件中解析图片
     * @param file xlsx 文件
     * @param sheetNum sheet 索引（从 0 开始）
     * @return 图片列表
     */
    public static List<ImageData> parse(File file, int sheetNum) {
        com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager manager = 
            com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager.getInstance();
        ExcelFileContext context = manager.getContext(file);
        return parse(context, sheetNum);
    }
    
    /**
     * 根据 sheet 名称解析图片
     * @param file xlsx 文件
     * @param sheetName sheet 名称
     * @return 图片列表
     */
    public static List<ImageData> parse(File file, String sheetName) {
        com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager manager = 
            com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager.getInstance();
        ExcelFileContext context = manager.getContext(file);
        return parse(context, sheetName);
    }
    
    /**
     * 根据行号和列号获取图片
     */
    public static ImageData getImage(File file, int sheetNum, int row, int column) {
        List<ImageData> images = parse(file, sheetNum);
        for (ImageData image : images) {
            if (image.getRow() == row && image.getColumn() == column) {
                return image;
            }
        }
        return null;
    }
    
    /**
     * 根据行号和列号获取图片（按 sheet 名称）
     */
    public static ImageData getImage(File file, String sheetName, int row, int column) {
        List<ImageData> images = parse(file, sheetName);
        for (ImageData image : images) {
            if (image.getRow() == row && image.getColumn() == column) {
                return image;
            }
        }
        return null;
    }
    
    /**
     * 根据行号获取该行的所有图片
     */
    public static List<ImageData> getImagesByRow(File file, int sheetNum, int row) {
        List<ImageData> images = parse(file, sheetNum);
        List<ImageData> result = new ArrayList<>();
        for (ImageData image : images) {
            if (image.getRow() == row) {
                result.add(image);
            }
        }
        return result;
    }
    
    /**
     * 根据行号获取该行的所有图片（按 sheet 名称）
     */
    public static List<ImageData> getImagesByRow(File file, String sheetName, int row) {
        List<ImageData> images = parse(file, sheetName);
        List<ImageData> result = new ArrayList<>();
        for (ImageData image : images) {
            if (image.getRow() == row) {
                result.add(image);
            }
        }
        return result;
    }
    
    // ==================== 内部解析方法 ====================
    
    /**
     * 解析 drawing.xml 获取图片引用信息
     */
    private static List<ImageRef> parseDrawingXml(PackagePart drawingPart) {
        List<ImageRef> refs = new ArrayList<>();
        
        try (InputStream is = drawingPart.getInputStream()) {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(is);
            
            // 解析两单元格锚点 (xdr:twoCellAnchor)
            NodeList anchors = doc.getElementsByTagNameNS("*", "twoCellAnchor");
            for (int i = 0; i < anchors.getLength(); i++) {
                Element anchor = (Element) anchors.item(i);
                ImageRef ref = parseAnchor(anchor, drawingPart, i);
                if (ref != null) {
                    refs.add(ref);
                }
            }
            
            // 解析绝对锚点 (xdr:absoluteAnchor)
            NodeList absoluteAnchors = doc.getElementsByTagNameNS("*", "absoluteAnchor");
            for (int i = 0; i < absoluteAnchors.getLength(); i++) {
                Element anchor = (Element) absoluteAnchors.item(i);
                ImageRef ref = parseAbsoluteAnchor(anchor, drawingPart, i);
                if (ref != null) {
                    refs.add(ref);
                }
            }
            
            // 解析单单元格锚点 (xdr:oneCellAnchor)
            NodeList oneCellAnchors = doc.getElementsByTagNameNS("*", "oneCellAnchor");
            for (int i = 0; i < oneCellAnchors.getLength(); i++) {
                Element anchor = (Element) oneCellAnchors.item(i);
                ImageRef ref = parseOneCellAnchor(anchor, drawingPart, i);
                if (ref != null) {
                    refs.add(ref);
                }
            }
            
        } catch (Exception e) {
            // 解析失败
        }
        
        return refs;
    }
    
    /**
     * 解析 twoCellAnchor
     */
    private static ImageRef parseAnchor(Element anchor, PackagePart drawingPart, int index) {
        try {
            // 获取起始单元格位置
            Element from = getChildElement(anchor, "from");
            if (from == null) return null;
            
            int row = getIntValue(getChildElement(from, "row"), 0);
            int col = getIntValue(getChildElement(from, "col"), 0);
            
            // 获取图片引用
            Element pic = getChildElement(anchor, "pic");
            if (pic == null) return null;
            
            Element blipFill = getChildElement(pic, "blipFill");
            if (blipFill == null) return null;
            
            Element blip = getChildElement(blipFill, "blip");
            if (blip == null) return null;
            
            String embed = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
            if (embed.isEmpty()) {
                embed = blip.getAttribute("embed");
            }
            if (embed.isEmpty()) {
                // 尝试获取 link 属性（外部链接）
                String link = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "link");
                if (link.isEmpty()) {
                    link = blip.getAttribute("link");
                }
                if (!link.isEmpty()) {
                    embed = link;
                }
            }
            
            ImageRef ref = new ImageRef();
            ref.row = row;
            ref.col = col;
            ref.relId = embed;
            ref.drawingPart = drawingPart;
            return ref;
            
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 解析 absoluteAnchor
     */
    private static ImageRef parseAbsoluteAnchor(Element anchor, PackagePart drawingPart, int index) {
        try {
            // 绝对锚点没有单元格引用，返回位置 0,0
            Element pic = getChildElement(anchor, "pic");
            if (pic == null) return null;
            
            Element blipFill = getChildElement(pic, "blipFill");
            if (blipFill == null) return null;
            
            Element blip = getChildElement(blipFill, "blip");
            if (blip == null) return null;
            
            String embed = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
            if (embed.isEmpty()) {
                embed = blip.getAttribute("embed");
            }
            
            ImageRef ref = new ImageRef();
            ref.row = 0;
            ref.col = index;
            ref.relId = embed;
            ref.drawingPart = drawingPart;
            return ref;
            
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 解析 oneCellAnchor
     */
    private static ImageRef parseOneCellAnchor(Element anchor, PackagePart drawingPart, int index) {
        try {
            Element from = getChildElement(anchor, "from");
            if (from == null) return null;
            
            int row = getIntValue(getChildElement(from, "row"), 0);
            int col = getIntValue(getChildElement(from, "col"), 0);
            
            Element pic = getChildElement(anchor, "pic");
            if (pic == null) return null;
            
            Element blipFill = getChildElement(pic, "blipFill");
            if (blipFill == null) return null;
            
            Element blip = getChildElement(blipFill, "blip");
            if (blip == null) return null;
            
            String embed = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
            if (embed.isEmpty()) {
                embed = blip.getAttribute("embed");
            }
            
            ImageRef ref = new ImageRef();
            ref.row = row;
            ref.col = col;
            ref.relId = embed;
            ref.drawingPart = drawingPart;
            return ref;
            
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 创建 ImageData 对象（使用上下文）
     */
    private static ImageData createImageData(ExcelFileContext context, ImageRef ref) {
        try {
            // 获取图片关系
            PackageRelationship imageRel = ref.drawingPart.getRelationship(ref.relId);
            if (imageRel == null) {
                return null;
            }
            
            // 获取图片 Part
            PackagePart imagePart = context.getPackage().getPart(imageRel);
            if (imagePart == null) {
                return null;
            }


            
            // 获取扩展名
            String contentType = imagePart.getContentType();
            String extension = MIME_TYPE_TO_EXTENSION.getOrDefault(contentType, "png");
            
            ImageData imageData = new ImageData();
            imageData.setRow(ref.row);
            imageData.setColumn(ref.col);
            imageData.setData(imagePart.getInputStream());
            imageData.setExtension(extension);
            imageData.setDispImg(false); // 浮动图片
            
            return imageData;
            
        } catch (Exception e) {
            return null;
        }
    }
    
    /**
     * 获取子元素
     */
    private static Element getChildElement(Element parent, String localName) {
        NodeList children = parent.getElementsByTagNameNS("*", localName);
        if (children.getLength() > 0) {
            return (Element) children.item(0);
        }
        return null;
    }
    
    /**
     * 获取整数值
     */
    private static int getIntValue(Element element, int defaultValue) {
        if (element == null) {
            return defaultValue;
        }
        try {
            String text = element.getTextContent();
            return Integer.parseInt(text.trim());
        } catch (Exception e) {
            return defaultValue;
        }
    }
    
    /**
     * 根据 MIME 类型获取文件扩展名
     */
    public static String getExtension(String mimeType) {
        return MIME_TYPE_TO_EXTENSION.getOrDefault(mimeType, "png");
    }
    
    /**
     * 图片引用信息（浮动图片）
     */
    private static class ImageRef {
        int row;
        int col;
        String relId;
        PackagePart drawingPart;
    }
    
    /**
     * 单元格图片引用（DISPIMG）
     */
    public record CellImageRef(int row,int col,String cellRef,String imageName,String formula) {

    }
}
