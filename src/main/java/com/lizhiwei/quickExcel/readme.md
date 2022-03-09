# quick-excel
## 更便捷的导入导出excel文件框架
### 使用说明
1. 在实体类的属性上加上注解@Excel value为导入时excel表头名称 name为导出时excel表头名称 index是导出时的排序 format为导出时的变更规则
2. 在需要导入的地方调用 ReadExcel类中的readExcel方法
3. 在需要导出的地方调用 DownloadExcel类中的setExcelProperty方法