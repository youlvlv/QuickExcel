package com.lizhiwei.quickExcel.v3.read.context;

import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilderFactory;
import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.util.*;

/**
 * Excel 文件上下文
 * 封装文件的所有元数据，避免重复解析
 * <p>
 * 设计原则：
 * 1. 只缓存全局元数据（共享字符串、workbook 结构）
 * 2. Sheet 内容按需加载，不缓存具体的 Sheet Part
 * 3. 支持并发读取同一个文件的不同 Sheet
 * </p>
 */
public class ExcelFileContext implements Closeable {
    
    private final File file;
    private final String fileKey;
    private final long lastModified;
    
    // OPC 包（延迟加载，需要时打开）
    private OPCPackage pkg;
    private boolean pkgOpened = false;
    
    // 共享字符串表（全局，一次性解析缓存）
    private SharedStrings sharedStrings;
    private List<String> sharedStringList;
    private boolean sharedStringsLoaded = false;
    
    // 样式表（全局，一次性解析缓存）
    private Styles styles;
    private boolean stylesLoaded = false;
    
    // Sheet 元数据（只缓存名称和关系 ID，不缓存 Part）
    private List<SheetMetadata> sheetMetadataList;
    private Map<String, Integer> sheetNameToIndex;
    private boolean sheetsMetadataLoaded = false;
    
    // Workbook Part（缓存）
    private PackagePart workbookPart;
    private PackageRelationship[] sheetRelationships;
    
    // DISPIMG 图片元数据缓存（cellimages.xml 和 cellimages.xml.rels）
    private Map<String, String> cellImagesNameToRid;
    private Map<String, String> cellImagesRidToTarget;
    private boolean cellImagesMetadataLoaded = false;

    private static final Logger log = LoggerFactory.getLogger(ExcelFileContext.class);
    
    /**
     * Sheet 元数据（轻量级，只包含名称和关系信息）
     */
    public static class SheetMetadata {
        public final int index;
        public final String name;
        public final String relId;
        public final int sheetId;
        
        public SheetMetadata(int index, String name, String relId, int sheetId) {
            this.index = index;
            this.name = name;
            this.relId = relId;
            this.sheetId = sheetId;
        }
    }
    
    /**
     * Sheet 信息（包含 Part，按需加载）
     */
    public static class SheetInfo {
        public final int index;
        public final String name;
        public final String relId;
        public final int sheetId;
        public final PackagePart part;
        
        public SheetInfo(SheetMetadata metadata, PackagePart part) {
            this.index = metadata.index;
            this.name = metadata.name;
            this.relId = metadata.relId;
            this.sheetId = metadata.sheetId;
            this.part = part;
        }
    }
    
    /**
     * 图片位置信息
     */
    public static class ImageLocation {
        public final int row;
        public final int col;
        public final String imageRelId;
        public final PackagePart drawingPart;
        
        public ImageLocation(int row, int col, String imageRelId, PackagePart drawingPart) {
            this.row = row;
            this.col = col;
            this.imageRelId = imageRelId;
            this.drawingPart = drawingPart;
        }
    }
    
    public ExcelFileContext(File file) {
        this.file = file;
        this.fileKey = file.getAbsolutePath();
        this.lastModified = file.lastModified();
        this.sheetMetadataList = new ArrayList<>();
        this.sheetNameToIndex = new HashMap<>();
    }
    
    /**
     * 获取文件
     */
    public File getFile() {
        return file;
    }
    
    /**
     * 获取文件键（用于缓存）
     */
    public String getFileKey() {
        return fileKey;
    }
    
    /**
     * 检查文件是否被修改
     */
    public boolean isModified() {
        return file.lastModified() != lastModified;
    }
    
    /**
     * 获取 OPCPackage（延迟加载）
     */
    public synchronized OPCPackage getPackage() throws Exception {
        if (!pkgOpened || pkg == null) {
            pkg = OPCPackage.open(file, PackageAccess.READ);
            pkgOpened = true;
        }
        return pkg;
    }
    
    /**
     * 获取共享字符串表
     */
    public synchronized SharedStrings getSharedStrings() throws Exception {
        if (!sharedStringsLoaded) {
            sharedStringList = parseSharedStrings(getPackage());
            sharedStrings = new SharedStringsWrapper(sharedStringList);
            sharedStringsLoaded = true;
        }
        return sharedStrings;
    }
    
    /**
     * 获取共享字符串列表（直接）
     */
    public synchronized List<String> getSharedStringList() throws Exception {
        if (!sharedStringsLoaded) {
            sharedStringList = parseSharedStrings(getPackage());
            sharedStrings = new SharedStringsWrapper(sharedStringList);
            sharedStringsLoaded = true;
        }
        return sharedStringList;
    }
    
    /**
     * 获取样式表
     */
    public synchronized Styles getStyles() throws Exception {
        if (!stylesLoaded) {
            styles = parseStyles(getPackage());
            stylesLoaded = true;
        }
        return styles;
    }
    
    /**
     * 获取 DISPIMG 图片元数据（图片名称 -> rId）
     */
    public synchronized Map<String, String> getCellImagesNameToRid() throws Exception {
        if (!cellImagesMetadataLoaded) {
            loadCellImagesMetadata();
        }
        return cellImagesNameToRid != null ? cellImagesNameToRid : Collections.emptyMap();
    }
    
    /**
     * 获取 DISPIMG 图片关系元数据（rId -> Target 路径）
     */
    public synchronized Map<String, String> getCellImagesRidToTarget() throws Exception {
        if (!cellImagesMetadataLoaded) {
            loadCellImagesMetadata();
        }
        return cellImagesRidToTarget != null ? cellImagesRidToTarget : Collections.emptyMap();
    }
    
    /**
     * 加载 cellimages.xml 和 cellimages.xml.rels 的元数据
     */
    private void loadCellImagesMetadata() throws Exception {
        cellImagesNameToRid = new HashMap<>();
        cellImagesRidToTarget = new HashMap<>();
        
        try {
            // 1. 解析 cellimages.xml.rels 获取 rId -> Target 映射
            PackagePart relsPart = null;
            for (PackagePart part : getPackage().getParts()) {
                String name = part.getPartName().getName();
                if (name.equals("/xl/_rels/cellimages.xml.rels") || name.endsWith("/cellimages.xml.rels")) {
                    relsPart = part;
                    break;
                }
            }
            
            if (relsPart != null) {
                try (var is = relsPart.getInputStream()) {
                    var factory = DocumentBuilderFactory.newInstance();
                    factory.setNamespaceAware(true);
                    var builder = factory.newDocumentBuilder();
                    var doc = builder.parse(is);
                    
                    var relationships = doc.getElementsByTagNameNS("*", "Relationship");
                    for (int i = 0; i < relationships.getLength(); i++) {
                        var rel = (Element) relationships.item(i);
                        String id = rel.getAttribute("Id");
                        String target = rel.getAttribute("Target");
                        if (!id.isEmpty() && !target.isEmpty()) {
                            cellImagesRidToTarget.put(id, target);
                        }
                    }
                }
            }
            
            // 2. 解析 cellimages.xml 获取图片名称 -> rId 映射
            PackagePart cellImagesPart = null;
            for (PackagePart part : getPackage().getParts()) {
                String name = part.getPartName().getName();
                if (name.equals("/xl/cellimages.xml") || name.endsWith("/cellimages.xml")) {
                    cellImagesPart = part;
                    break;
                }
            }
            
            if (cellImagesPart != null) {
                try (var is = cellImagesPart.getInputStream()) {
                    var factory = DocumentBuilderFactory.newInstance();
                    factory.setNamespaceAware(true);
                    var builder = factory.newDocumentBuilder();
                    var doc = builder.parse(is);
                    
                    // 查找所有 cellImage 元素
                    var cellImages = doc.getElementsByTagNameNS("*", "cellImage");
                    for (int i = 0; i < cellImages.getLength(); i++) {
                        var cellImage = (Element) cellImages.item(i);
                        
                        // 查找 cNvPr 元素获取 name 属性（图片名称）
                        String imageName = null;
                        var cnvs = cellImage.getElementsByTagNameNS("*", "cNvPr");
                        if (cnvs.getLength() > 0) {
                            var cnvPr = (Element) cnvs.item(0);
                            imageName = cnvPr.getAttribute("name");
                        }
                        
                        // 查找 blip 元素获取 embed 属性（rId）
                        String rid = null;
                        var blips = cellImage.getElementsByTagNameNS("*", "blip");
                        if (blips.getLength() > 0) {
                            var blip = (Element) blips.item(0);
                            rid = blip.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
                            if (rid == null || rid.isEmpty()) {
                                rid = blip.getAttribute("embed");
                            }
                        }
                        
                        if (imageName != null && !imageName.isEmpty() && rid != null && !rid.isEmpty()) {
                            cellImagesNameToRid.put(imageName, rid);
                        }
                    }
                }
            }
        } catch (Exception e) {
            // 解析失败，使用空映射
        }
        
        cellImagesMetadataLoaded = true;
    }
    
    /**
     * 获取所有 Sheet 元数据（轻量级，只包含名称和关系 ID）
     */
    public synchronized List<SheetMetadata> getSheetsMetadata() throws Exception {
        if (!sheetsMetadataLoaded) {
            parseSheetsMetadata();
        }
        return Collections.unmodifiableList(sheetMetadataList);
    }
    
    /**
     * 根据索引获取 Sheet 元数据
     */
    public SheetMetadata getSheetMetadata(int index) throws Exception {
        if (!sheetsMetadataLoaded) {
            parseSheetsMetadata();
        }
        if (index >= 0 && index < sheetMetadataList.size()) {
            return sheetMetadataList.get(index);
        }
        return null;
    }
    
    /**
     * 根据名称获取 Sheet 元数据
     */
    public SheetMetadata getSheetMetadata(String name) throws Exception {
        if (!sheetsMetadataLoaded) {
            parseSheetsMetadata();
        }
        Integer index = sheetNameToIndex.get(name);
        if (index != null && index >= 0 && index < sheetMetadataList.size()) {
            return sheetMetadataList.get(index);
        }
        return null;
    }
    
    /**
     * 根据索引获取 Sheet（按需加载 Part）
     */
    public SheetInfo getSheet(int index) throws Exception {
        SheetMetadata metadata = getSheetMetadata(index);
        if (metadata == null) {
            return null;
        }
        
        // 按需获取 Part
        PackagePart part = getSheetPart(metadata.relId);
        if (part == null) {
            return null;
        }
        
        return new SheetInfo(metadata, part);
    }
    
    /**
     * 根据名称获取 Sheet（按需加载 Part）
     */
    public SheetInfo getSheet(String name) throws Exception {
        SheetMetadata metadata = getSheetMetadata(name);
        if (metadata == null) {
            return null;
        }
        
        // 按需获取 Part
        PackagePart part = getSheetPart(metadata.relId);
        if (part == null) {
            return null;
        }
        
        return new SheetInfo(metadata, part);
    }
    
    /**
     * 获取 sheet 名称对应的索引
     */
    public int getSheetIndex(String name) throws Exception {
        if (!sheetsMetadataLoaded) {
            parseSheetsMetadata();
        }
        Integer index = sheetNameToIndex.get(name);
        return index != null ? index : -1;
    }
    
    /**
     * 获取 sheet 数量
     */
    public int getSheetCount() throws Exception {
        if (!sheetsMetadataLoaded) {
            parseSheetsMetadata();
        }
        return sheetMetadataList.size();
    }
    
    /**
     * 获取 Workbook Part
     */
    public synchronized PackagePart getWorkbookPart() throws Exception {
        if (workbookPart == null) {
            for (PackagePart part : getPackage().getParts()) {
                if ("/xl/workbook.xml".equals(part.getPartName().getName())) {
                    workbookPart = part;
                    break;
                }
            }
        }
        return workbookPart;
    }
    
    /**
     * 获取 Sheet Part（按需从关系获取）
     */
    private PackagePart getSheetPart(String relId) throws Exception {
        PackagePart wbPart = getWorkbookPart();
        if (wbPart == null) {
            log.error("[ExcelFileContext] workbookPart is null");
            return null;
        }
        
        PackageRelationship rel = wbPart.getRelationship(relId);
        if (rel == null) {
            log.error("[ExcelFileContext] Relationship not found for relId: {}", relId);
            return null;
        }
        
        log.info("[ExcelFileContext] Found relationship: {}",
                           ", target: " + rel.getTargetURI() + 
                           ", type: " + rel.getRelationshipType());
        
        // 尝试多种方式获取 Part
        
        // 方式 1：直接使用 getRelatedPart
        try {
            PackagePart part = wbPart.getRelatedPart(rel);
            if (part != null) {
                log.info("[ExcelFileContext] Got part via getRelatedPart: {}", part.getPartName().getName());
                return part;
            }
        } catch (Exception e) {
            log.error("[ExcelFileContext] getRelatedPart failed: {}", e.getMessage());
        }
        
        // 方式 2：手动解析目标 URI
        try {
            URI targetUri = rel.getTargetURI();
            String targetPath = targetUri.toString();
            
            // 处理相对路径
            if (!targetPath.startsWith("/")) {
                // 相对于 xl 目录
                targetPath = "/xl/" + targetPath;
            }
            
            log.info("[ExcelFileContext] Trying to get part with path: {}", targetPath);
            
            PackagePartName partName = PackagingURIHelper.createPartName(targetPath);
            PackagePart part = getPackage().getPart(partName);
            if (part != null) {
                log.info("[ExcelFileContext] Got part via manual path: {}", part.getPartName().getName());
                return part;
            }
        } catch (Exception e) {
            log.error("[ExcelFileContext] Manual path resolution failed: {}", e.getMessage());
        }
        
        // 方式 3：遍历所有 parts 查找匹配的
        try {
            String targetPath = rel.getTargetURI().toString();
            String expectedName = targetPath.contains("/") ? 
                targetPath.substring(targetPath.lastIndexOf('/') + 1) : targetPath;
            
            log.info("[ExcelFileContext] Searching for part with name: {}", expectedName);
            
            for (PackagePart part : getPackage().getParts()) {
                String partName = part.getPartName().getName();
                if (partName.endsWith(expectedName) || partName.contains(expectedName)) {
                    log.info("[ExcelFileContext] Found matching part: {}", partName);
                    return part;
                }
            }
        } catch (Exception e) {
            log.error("[ExcelFileContext] Part search failed: {}", e.getMessage());
        }
        
        log.error("[ExcelFileContext] Failed to get sheet part for relId: {}", relId);
        return null;
    }
    
    /**
     * 解析 Sheet 元数据（只解析名称和关系 ID，不获取 Part）
     */
    private void parseSheetsMetadata() throws Exception {
        PackagePart wbPart = getWorkbookPart();
        
        // 解析 workbook.xml 获取 sheet 名称和 relId
        Map<String, String> sheetNameMap = new HashMap<>();
        Map<String, String> sheetIdMap = new HashMap<>();
        
        try (var is = wbPart.getInputStream()) {
            var factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            var builder = factory.newDocumentBuilder();
            var doc = builder.parse(is);
            
            var sheetNodes = doc.getElementsByTagNameNS("*", "sheet");
            for (int i = 0; i < sheetNodes.getLength(); i++) {
                var sheetElement = (Element) sheetNodes.item(i);
                String name = sheetElement.getAttribute("name");
                String relId = sheetElement.getAttributeNS(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                if (relId.isEmpty()) {
                    relId = sheetElement.getAttribute("r:id");
                }
                String sheetIdStr = sheetElement.getAttribute("sheetId");
                
                sheetNameMap.put(relId, name);
                sheetIdMap.put(relId, sheetIdStr);
            }
        }
        
        // 只保存元数据，不获取 Part
        int index = 0;
        for (PackageRelationship rel : wbPart.getRelationships()) {
            if (rel.getRelationshipType().contains("worksheet")) {
                String relId = rel.getId();
                String name = sheetNameMap.getOrDefault(relId, "Sheet" + (index + 1));
                String sheetIdStr = sheetIdMap.getOrDefault(relId, String.valueOf(index + 1));
                int sheetId = Integer.parseInt(sheetIdStr);
                
                SheetMetadata metadata = new SheetMetadata(index, name, relId, sheetId);
                
                sheetMetadataList.add(metadata);
                sheetNameToIndex.put(name, index);
                index++;
            }
        }
        
        sheetsMetadataLoaded = true;
    }
    
    /**
     * 解析共享字符串
     */
    private List<String> parseSharedStrings(OPCPackage p) throws Exception {
        List<String> result = new ArrayList<>();
        
        // 查找 sharedStrings.xml
        PackagePart ssPart = null;
        for (PackagePart part : p.getParts()) {
            String name = part.getPartName().getName();
            if (name.contains("sharedStrings") || name.contains("SharedStrings")) {
                ssPart = part;
                break;
            }
        }
        
        if (ssPart == null) {
            return result;
        }
        
        try (var is = ssPart.getInputStream()) {
            var factory = javax.xml.parsers.DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            var builder = factory.newDocumentBuilder();
            var doc = builder.parse(is);
            
            var siNodes = doc.getElementsByTagNameNS("*", "si");
            for (int i = 0; i < siNodes.getLength(); i++) {
                var si = (org.w3c.dom.Element) siNodes.item(i);
                var tNodes = si.getElementsByTagNameNS("*", "t");
                if (tNodes.getLength() > 0) {
                    StringBuilder sb = new StringBuilder();
                    for (int j = 0; j < tNodes.getLength(); j++) {
                        sb.append(tNodes.item(j).getTextContent());
                    }
                    result.add(sb.toString());
                } else {
                    result.add("");
                }
            }
        }
        
        return result;
    }
    
    /**
     * 解析样式表
     */
    private Styles parseStyles(OPCPackage p) throws Exception {
        // 暂时返回 null，后续可以实现完整的样式解析
        return null;
    }
    
    @Override
    public void close() throws IOException {
        if (pkg != null && pkgOpened) {
            try {
                pkg.close();
            } catch (Exception e) {
                // 忽略
            }
            pkg = null;
            pkgOpened = false;
        }
    }
    
    /**
     * 清理缓存的数据（保留文件句柄）
     */
    public synchronized void clearCache() {
        sharedStringList = null;
        sharedStrings = null;
        sharedStringsLoaded = false;
        
        styles = null;
        stylesLoaded = false;
        
        sheetMetadataList.clear();
        sheetNameToIndex.clear();
        sheetsMetadataLoaded = false;
        
        cellImagesNameToRid = null;
        cellImagesRidToTarget = null;
        cellImagesMetadataLoaded = false;
    }
    
    /**
     * SharedStrings 包装器
     */
    private static class SharedStringsWrapper implements SharedStrings {
        private final List<String> strings;
        
        SharedStringsWrapper(List<String> strings) {
            this.strings = strings;
        }
        
        @Override
        public org.apache.poi.ss.usermodel.RichTextString getItemAt(int idx) {
            if (idx >= 0 && idx < strings.size()) {
                return new org.apache.poi.xssf.usermodel.XSSFRichTextString(strings.get(idx));
            }
            return null;
        }
        
        @Override
        public int getCount() {
            return strings.size();
        }
        
        @Override
        public int getUniqueCount() {
            return strings.size();
        }
    }
}
