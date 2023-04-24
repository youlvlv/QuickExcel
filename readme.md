<img alt="logo" src="/doc/image/logo.png" width="30%"/>

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

### 代码示例
1. 使用此框架中的方法生成一个Excl表格,在此表格中我们将完成但不限于以下几点 
- 自定义页眉
- 自定义页脚
- 合并单元格
- 自定义sheet名称
- 合并前后对比
<table>
    <tr>
        <td>第一列</td> 
        <td>第二列</td> 
   </tr>
    <tr>
        <td>这里是合并列</td>    
        <td >行一列二</td>  
    </tr>
    <tr>
        <td>这里是合并列</td>
        <td >行二列二</td>  
    </tr>
</table>
<hr>
<table>
    <tr>
        <td>第一列</td> 
        <td>第二列</td> 
   </tr>
    <tr>
        <td rowspan="2">这里是合并列</td>    
        <td >行一列二</td>  
    </tr>
    <tr>
        <td >行二列二</td>  
    </tr>
</table>

   ``` java
        String handMessage = "文字文字文字";
        // 将此集合根据指定字段分组        
        Map<String, List<T>> customerNameMap = newlist.stream().collect(Collectors.groupingBy(x -> x.getCustomerName()));
        // 创建Excl
        ExcelModel excelModel = DownloadComplexExcel.newExcel();
        // 输入当前sheet名称
        SheetModel sheetModel = excelModel.newSheet("sheet1")
        // 在当前sheet最上方插入一行，并且此行合并0-11列单元格插入文字 "我是合并"
                .newRow().setMergerValue(0, 11, "我是合并").over()
                // 在当前sheet最上方插入第二行，并且此行合并10-16列单元格插入文字为变量handMessage
                .newRow().setMergerValue(10, 16, handMessage).over()
                // 页眉创建结束
                .createHeader(TestTranslateExcl.class);
                // 此处以下是合并同类项的操作
        customerNameMap.forEach((k, v) -> {
            if (v.size() == 1) {
               // 如果当前集合在合并时出现长度为1的情况不要使用指定单元格合并操作
                sheetModel.createContent(TestTranslateExcl.class, v);
            } else {
               // 参数为 指定类映射,指定映射类的集合,new Since(所在列,"对应实体类名").......
                sheetModel.createContent(TestTranslateExcl.class, v, new Since(1, "customerName"), new Since(2, "customerMen"), new Since(3, "phone"), new Since(14, "contractDate"), new Since(0, "tableIndex"));
            }

        });
        // 添加页脚 在生成完数据之后的第一行行数进行操作:将低0-5列合并,并插入文字'合并',在第六列插入'第六列文字'.......已over()方法的调用结束当前行操作。
        sheetModel.newRow().setMergerValue(0, 5, "合计").setValue(6, "第六列文字").setValue(7, chejia[0] + "").over();
        // 导出Excl表格 注意 response类型为HttpServletResponse
        excelModel.exportExcelAndClose(DownloadComplexExcel.createDownload(response, "环卫车"));
   ```