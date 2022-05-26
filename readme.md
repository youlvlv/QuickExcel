# quick-excel
## 更便捷的导入导出excel文件框架
### 使用说明
---
#### 快捷使用
1. 在实体类的属性上加上注解@Excel value为导入时excel表头名称
name为导出时excel表头名称 index是导出时的排序 format为导出时的变更规则 topName为两行标题的首级标题
2. 在需要导入的地方调用 ReadExcel类中的readExcel方法
3. 在需要导出的地方调用 DownloadExcel类中的setExcelProperty方法
---
#### 进阶用法
1. 使用`DownloadComplexExcel.newExcel()`方法创建一个新的`ExcelModel`实例
2. 使用`ExcelModel`中的`newSheet()`方法创建新的`SheetModel`实例
3. `SheetModel`中分为 
- `createInfo()` 根据实体类导入信息
- `createHeader()` 根据实体类创建数据头
- `createContent()` 根据实体类列表创建数据内容
- `newRow()`创建新 `RowModel` 实例
- `newMoreRow()`创建 `MoreRowModel` 实例
4. `RowModel` 为excel中行 ,其中含有
- `setMergerValue()` 设置该行合并的单元格并设置值
- `setValue()` 设置该行指定单元格的值
5. `MoreRowModel` 为excel中相邻的两行，其中拥有方法
- `setValue()` 设置值
- `setMergerValue` 设置两个合并单元格的值