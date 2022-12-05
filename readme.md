<img alt="logo" src="doc/image/logo.png" width="50%"/>

# quick-excel

## 更便捷的导入导出excel文件框架

### 使用说明

当前版本: 2.0
---

#### 快捷使用

1. 在实体类的属性上加上注解`@Excel` value为导入时excel表头名称
   name为导出时excel表头名称 index是导出时的排序 format为导出时的变更规则 topName为两行标题的首级标题
2. 在需要导入的地方调用 `ReadExcel`类中的`readExcel`方法
3. 在需要导出的地方调用 `DownloadExcel`类中的`setExcelProperty`方法

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

### 引入方法

1. 使用maven的package命令 打包,从target中选择 QuickExcel-2.0-jar-with-dependencies.jar 文件 拷入需引入项目目录下，
   然后在引入项目的pom中填入如下内容
   ``` xml
   <dependency>
      <groupId>com.github</groupId>
      <artifactId>QuickExcel</artifactId>
      <version>2.0</version>
      <scope>system</scope>
      <systemPath>${pom.basedir}/jar包地址</systemPath>
    </dependency>
    ```
2. 使用maven的install命令,在需要引入的项目pom中填入如下内容
   ``` xml
     <dependency>
         <groupId>com.github</groupId>
         <artifactId>QuickExcel</artifactId>
         <version>2.0</version>
     </dependency>
   ```
 
