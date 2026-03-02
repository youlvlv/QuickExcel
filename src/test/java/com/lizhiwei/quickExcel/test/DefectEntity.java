package com.lizhiwei.quickExcel.test;

import com.lizhiwei.quickExcel.entity.Excel;

/**
 * 缺陷巡检实体类
 */
public class DefectEntity {

    /**
     * 线别
     */
    @Excel(value = "线别")
    private String lineType;

    /**
     * 区站
     */
    @Excel(value = "区站")
    private String station;

    /**
     * 行别
     */
    @Excel(value = "行别")
    private String rowType;

    /**
     * 杆号
     */
    @Excel(value = "杆号")
    private String poleNumber;

    /**
     * 缺陷类型
     */
    @Excel(value = "缺陷类型")
    private String defectType;

    /**
     * 发现时间
     */
    @Excel(value = "发现时间")
    private String discoveryTime;

    /**
     * 缺陷位置
     */
    @Excel(value = "缺陷位置")
    private String defectLocation;

    /**
     * 缺陷描述
     */
    @Excel(value = "缺陷描述")
    private String defectDescription;

    /**
     * 处理结果
     */
    @Excel(value = "处理结果")
    private String processingResult;

    /**
     * 处理前照片
     */
    @Excel(value = "处理前照片", isPicture = true)
    private String beforePhoto;

    /**
     * 处理后照片
     */
    @Excel(value = "处理后照片", isPicture = true)
    private String afterPhoto;

    /**
     * 缺陷来源
     */
    @Excel(value = "缺陷来源")
    private String defectSource;

    /**
     * 备注
     */
    @Excel(value = "备注")
    private String remarks;

    // Getters and Setters

    public String getLineType() {
        return lineType;
    }

    public void setLineType(String lineType) {
        this.lineType = lineType;
    }

    public String getStation() {
        return station;
    }

    public void setStation(String station) {
        this.station = station;
    }

    public String getRowType() {
        return rowType;
    }

    public void setRowType(String rowType) {
        this.rowType = rowType;
    }

    public String getPoleNumber() {
        return poleNumber;
    }

    public void setPoleNumber(String poleNumber) {
        this.poleNumber = poleNumber;
    }

    public String getDefectType() {
        return defectType;
    }

    public void setDefectType(String defectType) {
        this.defectType = defectType;
    }

    public String getDiscoveryTime() {
        return discoveryTime;
    }

    public void setDiscoveryTime(String discoveryTime) {
        this.discoveryTime = discoveryTime;
    }

    public String getDefectLocation() {
        return defectLocation;
    }

    public void setDefectLocation(String defectLocation) {
        this.defectLocation = defectLocation;
    }

    public String getDefectDescription() {
        return defectDescription;
    }

    public void setDefectDescription(String defectDescription) {
        this.defectDescription = defectDescription;
    }

    public String getProcessingResult() {
        return processingResult;
    }

    public void setProcessingResult(String processingResult) {
        this.processingResult = processingResult;
    }

    public String getBeforePhoto() {
        return beforePhoto;
    }

    public void setBeforePhoto(String beforePhoto) {
        this.beforePhoto = beforePhoto;
    }

    public String getAfterPhoto() {
        return afterPhoto;
    }

    public void setAfterPhoto(String afterPhoto) {
        this.afterPhoto = afterPhoto;
    }

    public String getDefectSource() {
        return defectSource;
    }

    public void setDefectSource(String defectSource) {
        this.defectSource = defectSource;
    }

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }
}
