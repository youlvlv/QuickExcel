package com.lizhiwei.quickExcel.v3.read.model;

import java.io.ByteArrayInputStream;
import java.io.InputStream;

/**
 * 图片数据模型
 */
public class ImageData {
    
    private int row;
    private int column;
    private String extension;
    private InputStream stream;
    
    /**
     * 图片名称（DISPIMG 函数中的图片标识）
     */
    private String imageName;
    
    /**
     * 是否为 DISPIMG 类型图片（WPS 嵌入单元格图片）
     */
    private boolean dispImg;
    
    /**
     * 原始公式（如果是 DISPIMG 类型）
     */
    private String formula;
    
    /**
     * 图片在 xlsx 中的路径
     */
    private String path;
    
    /**
     * 图片保存后的文件路径（由 ImageFileSaveFunction 返回）
     */
    private String savedPath;
    
    public ImageData() {
    }
    
    public ImageData(int row, int column, InputStream stream, String extension) {
        this.row = row;
        this.column = column;
        this.stream = stream;
        this.extension = extension;
    }
    
    /**
     * 创建 DISPIMG 类型的图片
     */
    public static ImageData createDispImg(int row, int column, String imageName, InputStream stream, String extension) {
        ImageData image = new ImageData(row, column, stream, extension);
        image.imageName = imageName;
        image.dispImg = true;
        return image;
    }
    
    public int getRow() {
        return row;
    }
    
    public void setRow(int row) {
        this.row = row;
    }
    
    public int getColumn() {
        return column;
    }
    
    public void setColumn(int column) {
        this.column = column;
    }
    
    /**
     * 获取图片数据流
     * <p>
     * 注意：每次调用都会创建新的 InputStream
     * </p>
     * 
     * @return InputStream，如果 bytes 为空返回 null
     */
    public InputStream getData() {
        if (stream == null) {
            return null;
        }
        return stream;
    }
    
    /**
     * 设置图片数据流
     * <p>
     * 注意：此方法会读取流中的所有字节到 bytes
     * </p>
     * 
     * @param data InputStream
     */
    public void setData(InputStream data) {
        if (data == null) {
            this.stream = null;
            return;
        }
    }
    
    public String getExtension() {
        return extension;
    }
    
    public void setExtension(String extension) {
        this.extension = extension;
    }

    
    public String getImageName() {
        return imageName;
    }
    
    public void setImageName(String imageName) {
        this.imageName = imageName;
    }
    
    public boolean isDispImg() {
        return dispImg;
    }
    
    public void setDispImg(boolean dispImg) {
        this.dispImg = dispImg;
    }
    
    public String getFormula() {
        return formula;
    }
    
    public void setFormula(String formula) {
        this.formula = formula;
    }
    
    public String getPath() {
        return path;
    }
    
    public void setPath(String path) {
        this.path = path;
    }
    
    /**
     * 获取图片保存后的文件路径
     */
    public String getSavedPath() {
        return savedPath;
    }
    
    /**
     * 设置图片保存后的文件路径
     */
    public void setSavedPath(String savedPath) {
        this.savedPath = savedPath;
    }
    
    /**
     * 获取图片的文件名（用于保存）
     */
    public String getFileName() {
        if (imageName != null && !imageName.isEmpty()) {
            return imageName + "." + extension;
        }
        return "image_" + row + "_" + column + "." + extension;
    }
    
    @Override
    public String toString() {
        return "ImageData{" +
                "row=" + row +
                ", column=" + column +
                ", extension='" + extension + '\'' +
                ", imageName='" + imageName + '\'' +
                ", dispImg=" + dispImg +
                ", savedPath='" + savedPath + '\'' +
                '}';
    }
}
