package com.lizhiwei.quickExcel.v2.core;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.PictureMap;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.ByteArrayInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

/**
 * 图片提取器
 * 负责从 Excel 中提取图片
 */
public class ImageExtractor {
    
    private static final Map<String, String> MIME_TYPE_TO_EXTENSION = new HashMap<>() {{
        put("image/jpeg", ".jpg");
        put("image/png", ".png");
        put("image/gif", ".gif");
        put("image/bmp", ".bmp");
    }};
    
    /**
     * 提取 Sheet 中的所有图片
     * @param sheet Sheet
     * @return 图片映射表
     */
    public static PictureMap extractPictures(Sheet sheet) {
        PictureMap pictureMap = new PictureMap();
        
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        Optional.ofNullable(drawing).ifPresent(draw -> {
            for (Object o : draw) {
                if (o instanceof Picture picture) {
                    ClientAnchor anchor = picture.getClientAnchor();
                    pictureMap.put(
                        new PictureMap.PictureKey(anchor.getRow1(), (int) anchor.getCol1()), 
                        picture
                    );
                }
            }
        });
        
        return pictureMap;
    }
    
    /**
     * 获取图片的文件扩展名
     */
    public static String getPictureExtension(Picture picture) {
        String mimeType = picture.getPictureData().getMimeType();
        return MIME_TYPE_TO_EXTENSION.getOrDefault(mimeType, "");
    }
    
    /**
     * 处理图片字段
     * @param pictureMap 图片映射
     * @param rowNum 行号
     * @param property 属性
     * @return 图片处理后的值
     */
    public static Object processImageField(PictureMap pictureMap, int rowNum, ExcelEntity property) {
        Picture picture = pictureMap.get(new PictureMap.PictureKey(rowNum, property.getValue()));
        if (picture == null) {
            return null;
        }
        
        ByteArrayInputStream inputStream = new ByteArrayInputStream(picture.getPictureData().getData());
        String extension = getPictureExtension(picture);
        
        return ExcelConfig.getImageFileFunction().apply(inputStream, extension);
    }
}
