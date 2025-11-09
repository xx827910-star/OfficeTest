# DOCX 格式提取工具

本项目包含使用 docx4j (Java) 和 python-docx (Python) 对 DOCX 文件进行全面格式提取的工具。

## Python 实现 (推荐)

Python 实现已经成功运行并提取了完整的格式信息。

### 使用方法

```bash
python3 docx_format_extractor.py
```

### 功能特性

提取的信息包括：
- 文档属性（作者、创建时间、修改时间等）
- 页面设置（尺寸、方向、边距）
- 样式定义（段落样式、字符样式、表格样式）
- 内容格式（段落对齐、缩进、间距、字体、字号、加粗/斜体等）
- 表格结构
- 编号和列表
- 图片数量
- 页眉和页脚

### 输出结果

运行结果已保存在 `test_docx_format_analysis.txt` 文件中。

## Java 实现 (docx4j)

Java 实现使用 docx4j 库。由于 Maven 网络/代理问题，采用手动下载依赖的方式。

### 当前状态

- ✅ Java 源代码已编写完成
- ✅ 代码编译成功
- ⚠️  运行时缺少部分依赖（org.docx4j.org.apache.xpath.XPathException）

### 已下载的依赖

```
lib/
├── activation.jar
├── commons-compress.jar
├── commons-io.jar
├── commons-lang3.jar
├── docx4j-core.jar
├── docx4j-ImportXHTML.jar
├── docx4j.jar
├── docx4j-openxml-objects.jar
├── istack-commons-runtime.jar
├── jakarta.xml.bind-api.jar
├── jaxb-core.jar
├── jaxb-runtime.jar
├── serializer.jar
├── slf4j-api.jar
├── slf4j-simple.jar
└── xalan.jar
```

### 编译

```bash
javac -cp "lib/*" -d build/classes src/main/java/com/docx4j/test/DocxFormatExtractor.java
```

### 运行

```bash
java -cp "lib/*:build/classes" com.docx4j.test.DocxFormatExtractor
```

### 遇到的问题

1. **Maven 网络问题**: 环境中的 Maven 无法正确通过代理连接到仓库
2. **依赖复杂性**: docx4j 11.4.9 有大量嵌套依赖，手动下载容易遗漏
3. **重新打包的类**: 某些类被 docx4j 重新打包（shaded），难以单独下载

## Test.docx 分析结果摘要

- **文件类型**: 学术论文（基于深度学习的图像识别技术研究与应用）
- **总段落数**: 237
- **表格数**: 5
- **图片数**: 2
- **样式数**: 164 (36个段落样式 + 27个字符样式 + 100个表格样式)
- **节数**: 6
- **页面尺寸**: 215.9mm × 279.4mm (A4)
- **页面方向**: 纵向

详细结果请查看 `test_docx_format_analysis.txt` 文件。
