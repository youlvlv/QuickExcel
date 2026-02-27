package com.lizhiwei.quickExcel.v3.read.util;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ImageFileSaveFunction;
import com.lizhiwei.quickExcel.v3.read.model.ImageData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.function.BiFunction;

/**
 * 图片保存工具类
 * 负责调用 ExcelConfig 或 ExcelEntity 中配置的图片保存函数
 */
public class ImageSaveUtil {

    private static final Logger logger = LoggerFactory.getLogger(ImageSaveUtil.class);

    
    /**
     * 保存图片并返回文件路径（使用 ExcelConfig 的全局配置）
     * 
     * @param imageData 图片数据
     * @return 保存后的文件路径，如果保存失败返回 null
     */
    public static String saveImage(ImageData imageData) {
        ImageFileSaveFunction saveFunction = ExcelConfig.getImageFileFunction();
        return saveImageWithFunction(imageData, saveFunction);
    }
    
    /**
     * 保存图片并返回文件路径（使用指定的保存函数）
     * 
     * @param imageData 图片数据
     * @param saveFunction 保存函数
     * @return 保存后的文件路径，如果保存失败返回 null
     */
    public static String saveImage(ImageData imageData, ImageFileSaveFunction saveFunction) {
        if (saveFunction == null) {
            return saveImage(imageData);
        }
        return saveImageWithFunction(imageData, saveFunction);
    }
    
    /**
     * 使用指定函数保存图片
     */
    private static String saveImageWithFunction(ImageData imageData, ImageFileSaveFunction saveFunction) {
        if (imageData == null) {
            return null;
        }
        
        // 如果已经保存过，直接返回已保存的路径
        if (imageData.getSavedPath() != null && !imageData.getSavedPath().isEmpty()) {
            return imageData.getSavedPath();
        }
        
        if (saveFunction == null) {
            return null;
        }
        
        // 获取 InputStream（支持 bytes 和 data 两种方式）
        try (InputStream inputStream = imageData.getData()) {
            if (inputStream == null) {
                return null;
            }
            
            String extension = imageData.getExtension();
            if (extension == null || extension.isEmpty()) {
                extension = "png";
            }
            // 确保扩展名以点开头
            if (!extension.startsWith(".")) {
                extension = "." + extension;
            }
            
            String savedPath = saveFunction.apply(inputStream, extension);
            
            // 保存路径到 ImageData
            if (savedPath != null && !savedPath.isEmpty()) {
                imageData.setSavedPath(savedPath);
            }
            
            return savedPath;
        } catch (Exception e) {
            logger.error("[ImageSaveUtil] Failed to save image: {}", e.getMessage());
            return null;
        }
    }
    
    /**
     * 为 ImageData 创建 InputStream
     * 
     * @param imageData 图片数据
     * @return InputStream，如果数据为空返回 null
     */
    public static InputStream createInputStream(ImageData imageData) {
        if (imageData == null) {
            return null;
        }
        return imageData.getData();
    }
    
    /**
     * 保存图片并更新 ImageData 对象
     * 
     * @param imageData 图片数据
     * @return 是否保存成功
     */
    public static boolean saveImageAndUpdate(ImageData imageData) {
        String path = saveImage(imageData);
        return path != null && !path.isEmpty();
    }
    
    /**
     * 保存图片并更新 ImageData 对象（使用指定的保存函数）
     * 
     * @param imageData 图片数据
     * @param saveFunction 保存函数
     * @return 是否保存成功
     */
    public static boolean saveImageAndUpdate(ImageData imageData, ImageFileSaveFunction saveFunction) {
        String path = saveImage(imageData, saveFunction);
        return path != null && !path.isEmpty();
    }
}
