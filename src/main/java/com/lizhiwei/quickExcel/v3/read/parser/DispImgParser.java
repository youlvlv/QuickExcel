package com.lizhiwei.quickExcel.v3.read.parser;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * DISPIMG 公式解析器
 * 用于解析 WPS Excel 中的 DISPIMG 函数
 * <p>
 * DISPIMG 函数格式：=DISPIMG("图片名称", 显示模式)
 * 示例：=DISPIMG("图片1", 1)
 * </p>
 */
public class DispImgParser {
    
    /**
     * DISPIMG 公式正则表达式
     * 匹配格式：=DISPIMG("图片名", 数字) 或 =DISPIMG('图片名', 数字)
     */
    private static final Pattern DISPIMG_PATTERN = Pattern.compile(
            "=DISPIMG\\s*\\(\\s*[\"']([^\"']+)[\"']\\s*,\\s*(\\d+)\\s*\\)",
            Pattern.CASE_INSENSITIVE
    );
    
    /**
     * 简化的 DISPIMG 检测模式
     */
    private static final Pattern SIMPLE_DISPIMG_PATTERN = Pattern.compile(
            "DISPIMG",
            Pattern.CASE_INSENSITIVE
    );
    
    /**
     * 解析结果
     */
    public static class DispImgInfo {
        /**
         * 图片名称/ID
         */
        private final String imageName;
        
        /**
         * 显示模式（通常为 1）
         */
        private final int displayMode;
        
        /**
         * 原始公式
         */
        private final String originalFormula;
        
        public DispImgInfo(String imageName, int displayMode, String originalFormula) {
            this.imageName = imageName;
            this.displayMode = displayMode;
            this.originalFormula = originalFormula;
        }
        
        public String getImageName() {
            return imageName;
        }
        
        public int getDisplayMode() {
            return displayMode;
        }
        
        public String getOriginalFormula() {
            return originalFormula;
        }
        
        @Override
        public String toString() {
            return "DispImgInfo{" +
                    "imageName='" + imageName + '\'' +
                    ", displayMode=" + displayMode +
                    '}';
        }
    }
    
    /**
     * 判断是否为 DISPIMG 公式
     *
     * @param formula 单元格公式或值
     * @return true 如果是 DISPIMG 公式
     */
    public static boolean isDispImgFormula(String formula) {
        if (formula == null || formula.isEmpty()) {
            return false;
        }
        return SIMPLE_DISPIMG_PATTERN.matcher(formula).find();
    }
    
    /**
     * 解析 DISPIMG 公式
     *
     * @param formula 单元格公式，如 =DISPIMG("图片1", 1)
     * @return DispImgInfo 对象，如果不是 DISPIMG 公式返回 null
     */
    public static DispImgInfo parse(String formula) {
        if (formula == null || formula.isEmpty()) {
            return null;
        }
        
        Matcher matcher = DISPIMG_PATTERN.matcher(formula);
        if (matcher.find()) {
            String imageName = matcher.group(1);
            int displayMode = Integer.parseInt(matcher.group(2));
            return new DispImgInfo(imageName, displayMode, formula);
        }
        
        return null;
    }
    
    /**
     * 从单元格值中提取图片名称
     *
     * @param cellValue 单元格值
     * @return 图片名称，如果不是 DISPIMG 公式返回 null
     */
    public static String extractImageName(String cellValue) {
        DispImgInfo info = parse(cellValue);
        return info != null ? info.getImageName() : null;
    }
    
    /**
     * 转换图片名称为文件搜索模式
     * WPS 存储图片时可能使用特定前缀或格式
     *
     * @param imageName 原始图片名称
     * @return 可能的文件名列表
     */
    public static String[] getPossibleFileNames(String imageName) {
        // WPS 可能使用的命名格式
        return new String[]{
                imageName,
                "DISPIMG_" + imageName,
                "image_" + imageName,
                // 添加常见扩展名前缀
                imageName + ".png",
                imageName + ".jpg",
                imageName + ".jpeg"
        };
    }
    
    /**
     * 判断是否为 WPS 生成的 xlsx 文件中的图片名称格式
     * WPS 通常使用 "_xlsInner_" 开头的名称
     *
     * @param imageName 图片名称
     * @return true 如果是 WPS 内部图片名称格式
     */
    public static boolean isWpsInternalName(String imageName) {
        return imageName != null && imageName.startsWith("_xlsInner_");
    }
}
