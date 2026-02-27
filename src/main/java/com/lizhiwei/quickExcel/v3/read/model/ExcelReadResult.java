package com.lizhiwei.quickExcel.v3.read.model;

import java.util.*;

/**
 * Excel 读取结果
 * 包含实体类列表和图片数据的完整读取结果
 *
 * @param <T> 实体类型
 */
public class ExcelReadResult<T> {
    
    /**
     * 实体类列表
     */
    private List<T> entities;
    
    /**
     * 所有图片列表（浮动图片 + DISPIMG 图片）
     */
    private List<ImageData> allImages;
    
    /**
     * 浮动图片列表（传统图片）
     */
    private List<ImageData> floatingImages;
    
    /**
     * DISPIMG 图片列表（WPS 嵌入单元格图片）
     */
    private List<ImageData> dispImgImages;
    
    /**
     * 按行索引分组的图片映射（key: 行索引, value: 该行的图片列表）
     */
    private Map<Integer, List<ImageData>> imagesByRow;
    
    /**
     * 按行和列索引定位的图片映射（key: "row,col", value: 图片）
     */
    private Map<String, ImageData> imageByPosition;
    
    /**
     * 按行和列索引定位的图片保存路径映射（key: "row,col", value: 保存后的文件路径）
     */
    private Map<String, String> imagePathByPosition;
    
    public ExcelReadResult() {
        this.entities = new ArrayList<>();
        this.allImages = new ArrayList<>();
        this.floatingImages = new ArrayList<>();
        this.dispImgImages = new ArrayList<>();
        this.imagesByRow = new HashMap<>();
        this.imageByPosition = new HashMap<>();
        this.imagePathByPosition = new HashMap<>();
    }
    
    /**
     * 添加实体
     */
    public void addEntity(T entity) {
        if (entity != null) {
            entities.add(entity);
        }
    }
    
    /**
     * 添加图片
     */
    public void addImage(ImageData image) {
        if (image == null) {
            return;
        }
        allImages.add(image);
        
        // 根据图片类型分类
        if (image.isDispImg()) {
            dispImgImages.add(image);
        } else {
            floatingImages.add(image);
        }
        
        // 添加到行映射
        int row = image.getRow();
        imagesByRow.computeIfAbsent(row, k -> new ArrayList<>()).add(image);
        
        // 添加到位置映射
        String key = buildPositionKey(image.getRow(), image.getColumn());
        imageByPosition.put(key, image);
        
        // 如果有保存路径，添加到路径映射
        if (image.getSavedPath() != null && !image.getSavedPath().isEmpty()) {
            imagePathByPosition.put(key, image.getSavedPath());
        }
    }
    
    /**
     * 批量添加图片
     */
    public void addImages(List<ImageData> images) {
        if (images == null) {
            return;
        }
        for (ImageData image : images) {
            addImage(image);
        }
    }
    
    /**
     * 根据行号和列号获取图片
     */
    public ImageData getImage(int row, int column) {
        return imageByPosition.get(buildPositionKey(row, column));
    }
    
    /**
     * 根据行号和列号获取图片保存后的路径
     */
    public String getImagePath(int row, int column) {
        return imagePathByPosition.get(buildPositionKey(row, column));
    }
    
    /**
     * 根据行号获取该行的所有图片
     */
    public List<ImageData> getImagesByRow(int row) {
        return imagesByRow.getOrDefault(row, Collections.emptyList());
    }
    
    /**
     * 根据行号获取该行所有图片的保存路径
     */
    public List<String> getImagePathsByRow(int row) {
        List<ImageData> images = imagesByRow.getOrDefault(row, Collections.emptyList());
        List<String> paths = new ArrayList<>();
        for (ImageData image : images) {
            if (image.getSavedPath() != null && !image.getSavedPath().isEmpty()) {
                paths.add(image.getSavedPath());
            }
        }
        return paths;
    }
    
    /**
     * 根据行号和列号范围获取图片列表
     */
    public List<ImageData> getImagesInRange(int startRow, int endRow, int startCol, int endCol) {
        List<ImageData> result = new ArrayList<>();
        for (ImageData image : allImages) {
            int row = image.getRow();
            int col = image.getColumn();
            if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) {
                result.add(image);
            }
        }
        return result;
    }
    
    /**
     * 获取指定实体行对应的图片
     * 注意：实体行索引 = 数据行索引 - startRow
     */
    public List<ImageData> getImagesForEntity(int entityIndex, int startRow) {
        int actualRow = startRow + entityIndex;
        return getImagesByRow(actualRow);
    }
    
    /**
     * 获取指定实体行对应的图片保存路径
     */
    public List<String> getImagePathsForEntity(int entityIndex, int startRow) {
        int actualRow = startRow + entityIndex;
        return getImagePathsByRow(actualRow);
    }
    
    /**
     * 获取所有已保存的图片路径
     */
    public List<String> getAllSavedPaths() {
        List<String> paths = new ArrayList<>();
        for (ImageData image : allImages) {
            if (image.getSavedPath() != null && !image.getSavedPath().isEmpty()) {
                paths.add(image.getSavedPath());
            }
        }
        return paths;
    }
    
    /**
     * 构建位置键
     */
    private String buildPositionKey(int row, int column) {
        return row + "," + column;
    }
    
    // ==================== Getters ====================
    
    public List<T> getEntities() {
        return entities;
    }
    
    public void setEntities(List<T> entities) {
        this.entities = entities != null ? entities : new ArrayList<>();
    }
    
    public List<ImageData> getAllImages() {
        return allImages;
    }
    
    public List<ImageData> getFloatingImages() {
        return floatingImages;
    }
    
    public List<ImageData> getDispImgImages() {
        return dispImgImages;
    }
    
    public Map<Integer, List<ImageData>> getImagesByRow() {
        return imagesByRow;
    }
    
    public Map<String, String> getImagePathByPosition() {
        return imagePathByPosition;
    }
    
    /**
     * 获取实体数量
     */
    public int getEntityCount() {
        return entities.size();
    }
    
    /**
     * 获取图片总数
     */
    public int getImageCount() {
        return allImages.size();
    }
    
    /**
     * 获取已保存图片数量
     */
    public int getSavedImageCount() {
        int count = 0;
        for (ImageData image : allImages) {
            if (image.getSavedPath() != null && !image.getSavedPath().isEmpty()) {
                count++;
            }
        }
        return count;
    }
    
    /**
     * 是否包含图片
     */
    public boolean hasImages() {
        return !allImages.isEmpty();
    }
    
    /**
     * 是否包含实体数据
     */
    public boolean hasEntities() {
        return !entities.isEmpty();
    }
    
    @Override
    public String toString() {
        return "ExcelReadResult{" +
                "entities=" + entities.size() +
                ", allImages=" + allImages.size() +
                ", floatingImages=" + floatingImages.size() +
                ", dispImgImages=" + dispImgImages.size() +
                ", savedImages=" + getSavedImageCount() +
                '}';
    }
}
