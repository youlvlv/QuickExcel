# ReadExcel 使用指南

## 概述

`ReadExcel` 现在支持根据配置自动选择 V2（Apache POI）或 V3（SAX/DOM）解析引擎。

- **V2 引擎（默认）**：基于 Apache POI，功能完整，兼容性好
- **V3 引擎**：基于 SAX/DOM 解析，内存占用更低，适合大文件

## 配置方式

### 全局配置引擎

```java
import com.lizhiwei.quickExcel.config.ExcelConfig;

// 设置为 V3 引擎（全局生效）
ExcelConfig.setDefaultReadEngine(ExcelConfig.ReadEngine.V3);

// 恢复为 V2 引擎（默认）
ExcelConfig.setDefaultReadEngine(ExcelConfig.ReadEngine.V2);
```

### 直接使用特定引擎

```java
import com.lizhiwei.quickExcel.util.ReadExcel;

// 使用 V2 引擎
List<User> users = ReadExcel.readExcelByV2(file, 1, 0, 0, User.class);

// 使用 V3 引擎
List<User> users = ReadExcel.readExcelByV3(file, 1, 0, 0, User.class);
```

## 基础用法

### 简单读取

```java
File file = new File("/path/to/excel.xlsx");
List<User> users = ReadExcel.readExcel(file, 1, 0, 0, User.class);
```

### 带 safe 模式（收集所有错误）

```java
// safe = true：收集所有行错误后统一抛出 ExcelReadException
try {
    List<User> users = ReadExcel.readExcel(file, 1, 0, 0, User.class, true);
} catch (ExcelReadException e) {
    List<ReadErrorInfo> errors = e.getErrorInfos();
    // 处理错误列表
}
```

### 读取图片

```java
List<User> users = ReadExcel.readExcel(file, 1, 0, 0, User.class, false, true);
```

### 读取为 Map 列表

```java
List<Map<String, String>> data = ReadExcel.readExcel(
    file, 1, 0, 0, false, ReadExcel.getExcelEntities(User.class)
);
```

## V2 引擎拆分后的组件

V2 引擎已拆分为职责单一的组件：

| 组件 | 职责 |
|------|------|
| `WorkbookLoader` | 加载 Excel 文件为 Workbook |
| `CellValueExtractor` | 提取单元格各种类型的值 |
| `MergeCellResolver` | 处理合并单元格 |
| `HeaderMatcher` | 匹配表头与实体属性 |
| `ImageExtractor` | 提取图片 |
| `EntityValueConverter` | 转换值为实体字段类型 |
| `EntityPopulator` | 填充值到实体对象 |
| `EmptyRowChecker` | 检查空行 |

### 直接使用 V2 组件

```java
import com.lizhiwei.quickExcel.v2.reader.PoiExcelReader;
import com.lizhiwei.quickExcel.core.ExcelReader;

ExcelReader reader = new PoiExcelReader();
List<User> users = reader.readExcel(file, 1, 0, 0, User.class);
```

## V3 引擎用法

### 基础读取

```java
List<User> users = ReadExcel.readExcelByV3(file, 1, 0, 0, User.class);
```

### 指定解析策略

```java
import com.lizhiwei.quickExcel.v3.read.ReadExcelByXml;

// SAX 模式（默认，内存占用最低）
List<User> users = ReadExcel.readExcelByV3(
    file, 1, 0, 0, User.class, false, ReadExcelByXml.ReadStrategy.SAX
);

// DOM 模式（速度更快，但内存占用较高）
List<User> users = ReadExcel.readExcelByV3(
    file, 1, 0, 0, User.class, false, ReadExcelByXml.ReadStrategy.DOM
);
```

### 读取图片

```java
List<ImageData> images = ReadExcel.readExcelImagesByV3(file, 0);
```

## 引擎选择建议

| 场景 | 推荐引擎 | 原因 |
|------|----------|------|
| 小文件（< 10MB） | V2 | 功能完整，兼容性好 |
| 大文件（> 10MB） | V3 (SAX) | 内存占用低，不会 OOM |
| 需要复杂格式化 | V2 | 支持更多格式和公式 |
| 只需要原始数据 | V3 | 解析速度快 |
| 包含大量图片 | V2 | 图片处理更完善 |

## 兼容性说明

- 旧版 `ReadExcel.readExcelByPoi()` 和 `ReadExcel.readExcelByXml()` 方法仍然可用，但已标记为 `@Deprecated`
- 建议迁移到新的 `readExcelByV2()` 和 `readExcelByV3()` 方法
