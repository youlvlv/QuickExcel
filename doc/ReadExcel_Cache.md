# V3 引擎文件缓存机制

## 概述

V3 引擎使用**单例上下文管理器**缓存文件元数据，避免重复解析同一个文件。

**核心设计原则：**
1. 只缓存全局元数据（共享字符串、workbook 结构）
2. **Sheet 内容按需加载**，不缓存具体的 Sheet Part
3. 支持并发读取同一个文件的不同 Sheet

## 核心组件

### ExcelFileContext（文件上下文）

#### 缓存的内容（全局，一次性解析）
- **共享字符串表** - 整个文件共享，解析一次
- **Workbook 结构** - Sheet 名称和索引的映射
- **样式表** - 全局样式定义

#### 不缓存的内容（按需加载）
- **Sheet Part** - 只在读取具体 Sheet 时获取
- **Sheet 内容流** - 流式读取，不驻留内存
- **图片数据** - 按需解析

### ExcelFileContextManager（上下文管理器）
- 单例模式管理所有文件上下文
- 基于文件路径 + 修改时间作为缓存键
- 自动检测文件修改并刷新缓存
- 支持弱引用，内存紧张时自动释放

## 内存占用分析

### 传统方式（每次重新解析）
```java
// 假设文件有 10 个 sheet，每个 sheet 1MB
for (int i = 0; i < 10; i++) {
    List<User> users = ReadExcelByXml.readExcel(file, 1, 0, i, User.class);
}
// 共打开文件 10 次
// 共解析共享字符串 10 次
// 共解析 workbook.xml 10 次
// 内存占用波动大
```

### 缓存方式（优化后）
```java
// 假设文件有 10 个 sheet，共享字符串 100KB
for (int i = 0; i < 10; i++) {
    List<User> users = ReadExcelByXml.readExcel(file, 1, 0, i, User.class);
}
// 打开文件 1 次
// 解析共享字符串 1 次（常驻内存 100KB）
// 解析 workbook.xml 1 次（元数据很小）
// 每个 sheet 按需读取，流式处理，不累积内存
```

## 使用方法

### 自动缓存（默认行为）
```java
// 第一次读取 - 创建上下文，解析共享字符串和 workbook 结构
List<User> users1 = ReadExcel.readExcelByV3(file, 1, 0, 0, User.class);

// 第二次读取同一文件的不同 sheet - 复用上下文，只加载指定 sheet
List<User> users2 = ReadExcel.readExcelByV3(file, 1, 0, 1, User.class);

// 读取图片 - 同样复用上下文
List<ImageData> images = ReadExcel.readExcelImagesByV3(file, 0);
```

### 手动管理上下文
```java
import com.lizhiwei.quickExcel.v3.read.ReadExcelByXml;
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContext;

// 获取文件上下文
ExcelFileContext context = ReadExcelByXml.getContext(file);

// 获取所有 sheet 名称（只返回元数据，不加载内容）
List<String> sheetNames = ReadExcelByXml.getSheetNames(file);

// 获取 sheet 名称对应的索引
int index = ReadExcelByXml.getSheetIndex(file, "Sheet1");

// 获取索引对应的 sheet 名称
String name = ReadExcelByXml.getSheetName(file, 0);

// 按需获取指定 sheet 的信息（此时才加载 Part）
ExcelFileContext.SheetInfo sheetInfo = context.getSheet(0);

// 清理指定文件的缓存
ReadExcelByXml.clearContext(file);

// 清理所有缓存
ReadExcelByXml.clearAllContexts();
```

### 上下文管理器配置
```java
import com.lizhiwei.quickExcel.v3.read.context.ExcelFileContextManager;

ExcelFileContextManager manager = ExcelFileContextManager.getInstance();

// 是否使用弱引用（默认 true）
// 内存紧张时允许 GC 回收缓存
manager.setUseWeakReference(true);

// 是否自动清理被修改的文件缓存（默认 true）
manager.setAutoCleanModified(true);

// 获取当前缓存数量
int size = manager.getCacheSize();

// 手动清理过期缓存
manager.cleanUp();

// 清理所有缓存
manager.clearAll();
```

## API 对比

### 获取 Sheet 元数据（轻量级）
```java
// 只获取名称和索引信息，不加载内容
ExcelFileContext.SheetMetadata metadata = context.getSheetMetadata(0);
System.out.println(metadata.name);  // "Sheet1"
System.out.println(metadata.index); // 0
```

### 获取 Sheet 完整信息（按需加载）
```java
// 获取元数据并加载 Part（用于读取内容）
ExcelFileContext.SheetInfo sheetInfo = context.getSheet(0);
InputStream is = sheetInfo.part.getInputStream();  // 流式读取
```

## 线程安全

- `ExcelFileContextManager` 使用 `ConcurrentHashMap` 保证线程安全
- 多个线程可以同时读取同一个文件的不同 Sheet
- 共享字符串等全局数据线程安全（只读）

## 注意事项

1. **Sheet 内容不缓存**：每次读取 Sheet 都会重新打开流，但 OPC Package 复用
2. **大文件友好**：即使有 100 个 sheet，也只会缓存一份共享字符串和元数据
3. **流式读取**：Sheet 数据使用 SAX/DOM 流式解析，不会全部加载到内存
4. **文件锁**：OPC Package 打开期间文件不会被锁定，但修改后需要清理缓存

## 性能建议

### 批量读取多个 Sheet
```java
// 推荐：复用上下文
ExcelFileContext context = ReadExcelByXml.getContext(file);
for (int i = 0; i < context.getSheetCount(); i++) {
    List<User> users = ReadExcel.readExcelByV3(file, 1, 0, i, User.class);
}
```

### 只获取 Sheet 名称（不读取内容）
```java
// 轻量级操作，只解析 workbook.xml
List<String> names = ReadExcelByXml.getSheetNames(file);
```
