# Word 文档格式提取器 - 增强版

## 📖 简介

这是一个基于 **Open XML SDK 3.x** 开发的 Word 文档格式完整提取工具，参考 [Microsoft 官方文档](https://learn.microsoft.com/en-us/office/open-xml/) 和社区最佳实践实现。

### ✨ 主要特性

- ✅ **完整提取**: 提取文档的所有格式信息（237个段落 vs 云端版20个）
- ✅ **双格式输出**: 同时生成 TXT（人类可读）和 JSON（机器可解析）
- ✅ **图片支持**: 提取图片的类型、大小、尺寸、名称
- ✅ **页眉页脚**: 提取所有页眉和页脚内容
- ✅ **超链接和书签**: 提取文档中的所有超链接和书签
- ✅ **主题和批注**: 支持主题信息和批注提取
- ✅ **扩展属性**: 提取应用程序、字数统计等扩展属性

## 🚀 快速开始

### 环境要求

- .NET SDK 8.0+
- DocumentFormat.OpenXml 3.3.0

### 安装

```bash
# 克隆或下载项目
cd DocxFormatExtractor

# 还原依赖
dotnet restore

# 编译
dotnet build
```

### 运行

```bash
# 运行程序（默认输出TXT和JSON两种格式）
dotnet run

# 输出文件
# - format_output_enhanced.txt  (5.6KB, 人类可读)
# - format_output_enhanced.json (431KB, 完整数据)
```

### 自定义输入文件

修改 `EnhancedProgram.cs` 中的路径：

```csharp
string docPath = "/path/to/your/document.docx";
```

### 自定义输出格式

```csharp
// 仅输出TXT
string outputFormat = "txt";

// 仅输出JSON
string outputFormat = "json";

// 同时输出（默认）
string outputFormat = "both";
```

## 📊 提取内容详解

### 1. 文档属性

```json
{
  "Title": "",
  "Creator": "python-docx",
  "Application": "Microsoft Macintosh Word",
  "Pages": "1",
  "Words": "0",
  "Characters": "0",
  "Revision": "1",
  ...
}
```

**包含信息**:
- 基本属性：标题、主题、创建者、关键词
- 扩展属性：应用程序、公司、页数、字数、字符数
- 文档设置：缩放比例、默认制表位

### 2. 样式信息（164个样式）

```json
{
  "StyleId": "Heading1",
  "StyleName": "heading 1",
  "Type": "paragraph",
  "BasedOn": "Normal",
  "ParagraphProperties": {
    "Alignment": "center",
    "SpacingBefore": "480"
  },
  "RunProperties": {
    "FontSize": "28",
    "Bold": true,
    "Color": "365F91"
  }
}
```

**包含信息**:
- 样式ID和名称
- 样式类型和继承关系
- 段落属性（对齐、缩进、间距）
- 文本属性（字体、字号、颜色、粗体、斜体）

### 3. 段落和文本（237个段落）

```json
{
  "Index": 0,
  "Text": "摘  要",
  "StyleId": "",
  "Alignment": "center",
  "SpacingBefore": "240",
  "Runs": [
    {
      "Text": "摘  要",
      "FontNameAscii": "宋体",
      "FontNameEastAsia": "宋体",
      "FontSize": "32",
      "Bold": true,
      "Color": "000000"
    }
  ]
}
```

**包含信息**:
- 段落索引和文本内容
- 样式ID、对齐方式
- 缩进（左、右、首行、悬挂）
- 间距（段前、段后、行距）
- 编号属性
- 边框和底纹
- 每个文本运行的详细格式

### 4. 表格（5个表格）

```json
{
  "Index": 0,
  "StyleId": "TableNormal",
  "Width": "0",
  "HasBorders": true,
  "Rows": [
    {
      "Height": "400",
      "IsHeader": true,
      "Cells": [
        {
          "Text": "网络名称",
          "Width": "1615",
          "BackgroundColor": "FFFFFF"
        }
      ]
    }
  ]
}
```

**包含信息**:
- 表格样式、宽度、对齐
- 边框信息
- 每行的高度、是否标题行
- 每个单元格的文本、宽度、背景色、对齐方式、合并信息

### 5. 图片（2张图片）

```json
{
  "Index": 0,
  "ContentType": "image/png",
  "RelationshipId": "rId13",
  "SizeBytes": 2377,
  "Width": "4572000",
  "Height": "2286000",
  "Name": "Picture 1",
  "Description": ""
}
```

**包含信息**:
- 图片类型（PNG/JPEG等）
- 关系ID
- 文件大小（字节）
- 尺寸（宽度×高度，EMU单位）
- 图片名称和描述

### 6. 节信息（1个节）

```json
{
  "Index": 0,
  "PageWidth": "12240",
  "PageHeight": "15840",
  "Orientation": "Portrait",
  "MarginTop": "1134",
  "MarginBottom": "1134",
  "MarginLeft": "1417",
  "MarginRight": "1134",
  "ColumnCount": "1"
}
```

**包含信息**:
- 页面尺寸和方向
- 页边距（上下左右、页眉页脚距离、装订线）
- 分栏信息

### 7. 页眉和页脚（4页眉 + 3页脚）

```json
{
  "Index": 0,
  "Text": "基于深度学习的图像识别技术研究与应用",
  "RelationshipId": "rId7"
}
```

### 8. 超链接（68个）

```json
{
  "Index": 0,
  "Text": "第1章 绪论",
  "Url": "",
  "Anchor": "_Chapter_1",
  "IsExternal": false
}
```

**包含信息**:
- 超链接文本
- 目标URL（外部链接）
- 锚点（内部链接）
- 是否外部链接

### 9. 书签（68个）

```json
{
  "Index": 0,
  "Id": "1",
  "Name": "_Chapter_1"
}
```

### 10. 字体表（8种字体）

```json
{
  "Name": "Times New Roman",
  "Family": "roman",
  "Pitch": "variable"
}
```

### 11. 编号系统（9个编号定义）

```json
{
  "AbstractNumId": "0",
  "LevelCount": 1,
  "Levels": [
    {
      "LevelIndex": "0",
      "NumberFormat": "decimal",
      "LevelText": "%1.",
      "StartValue": "1"
    }
  ]
}
```

### 12. 主题和批注

```json
{
  "ThemeName": "Office Theme",
  "Comments": []
}
```

## 🎯 使用场景

### 场景1: 文档格式审计

查看TXT报告快速了解文档结构：

```bash
cat format_output_enhanced.txt
```

输出示例：
```
【1. 文档属性】
标题:
创建者: python-docx
应用程序: Microsoft Macintosh Word
页数: 1

【3. 段落】 总数: 237
段落 #0: 摘  要
段落 #2: 随着深度学习技术的快速发展...

【5. 图片】 总数: 2
图片 #0: image/png, 2377字节

【7. 超链接】 总数: 68
```

### 场景2: 程序化处理

使用 jq 工具处理JSON数据：

```bash
# 查看所有段落
cat format_output_enhanced.json | jq '.Paragraphs'

# 统计段落数
cat format_output_enhanced.json | jq '.Paragraphs | length'
# 输出: 237

# 提取所有图片信息
cat format_output_enhanced.json | jq '.Images[]'

# 查找包含特定文本的段落
cat format_output_enhanced.json | jq '.Paragraphs[] | select(.Text | contains("深度学习"))'

# 提取所有外部超链接
cat format_output_enhanced.json | jq '.Hyperlinks[] | select(.IsExternal == true)'

# 统计样式数量
cat format_output_enhanced.json | jq '.Styles | length'
# 输出: 164

# 查看文档属性
cat format_output_enhanced.json | jq '.DocumentProperties'
```

### 场景3: 图片资源管理

```bash
# 提取图片列表
cat format_output_enhanced.json | jq '.Images[] | {Name, ContentType, Size: .SizeBytes}'

# 输出:
# {
#   "Name": "Picture 1",
#   "ContentType": "image/png",
#   "Size": 2377
# }
```

### 场景4: 样式分析

```bash
# 查找使用了颜色的样式
cat format_output_enhanced.json | jq '.Styles[] | select(.RunProperties.Color != "")'

# 查找粗体样式
cat format_output_enhanced.json | jq '.Styles[] | select(.RunProperties.Bold == true) | .StyleName'
```

### 场景5: 表格数据提取

```bash
# 提取第一个表格的所有单元格文本
cat format_output_enhanced.json | jq '.Tables[0].Rows[].Cells[].Text'
```

## 📚 代码结构

```
DocxFormatExtractor/
├── EnhancedProgram.cs          # 主程序（增强版）
├── Program.cs                  # 原始版本
├── DocxFormatExtractor.csproj  # 项目文件
└── README.md                   # 本文档
```

### 核心类

```csharp
// 主数据模型
public class DocumentFormatInfo
{
    public DocumentPropertiesInfo DocumentProperties { get; set; }
    public List<StyleInfo> Styles { get; set; }
    public List<ParagraphInfo> Paragraphs { get; set; }
    public List<TableInfo> Tables { get; set; }
    public List<SectionInfo> Sections { get; set; }
    public List<ImageInfo> Images { get; set; }
    public List<HeaderFooterInfo> Headers { get; set; }
    public List<HeaderFooterInfo> Footers { get; set; }
    public List<HyperlinkInfo> Hyperlinks { get; set; }
    public List<BookmarkInfo> Bookmarks { get; set; }
    public List<FontInfo> Fonts { get; set; }
    public List<NumberingInfo> Numbering { get; set; }
    public List<CommentInfo> Comments { get; set; }
    public string ThemeName { get; set; }
}
```

### 提取方法

```csharp
ExtractDocumentProperties()    // 文档属性
ExtractStyles()                // 样式
ExtractParagraphsAndRuns()     // 段落和文本
ExtractTables()                // 表格
ExtractSections()              // 节
ExtractImages()                // 图片
ExtractHeadersFooters()        // 页眉页脚
ExtractHyperlinksAndBookmarks() // 超链接和书签
ExtractFontsAndNumbering()     // 字体和编号
ExtractThemesAndComments()     // 主题和批注
```

## 🔧 高级配置

### 自定义提取深度

在代码中注释掉不需要的提取模块：

```csharp
// ExtractImages(doc);           // 跳过图片提取
// ExtractHeadersFooters(doc);   // 跳过页眉页脚
```

### 性能优化

对于超大文档，可以限制段落提取：

```csharp
// 仅提取前100个段落
foreach (var para in body.Elements<Paragraph>().Take(100))
```

### JSON格式化选项

```csharp
var options = new JsonSerializerOptions
{
    WriteIndented = true,  // 格式化输出
    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};
```

## 📖 参考文档

- [Open XML SDK 官方文档](https://learn.microsoft.com/en-us/office/open-xml/)
- [GitHub 仓库](https://github.com/dotnet/Open-XML-SDK)
- [NuGet 包](https://www.nuget.org/packages/DocumentFormat.OpenXml/)
- [图片提取示例](https://pinkhatcode.com/2017/09/01/extract-images-word-document-using-openxml-c/)

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可

MIT License

## 🙏 致谢

- Microsoft Open XML SDK 团队
- .NET 社区
- 所有贡献者

---

**版本**: 1.0.0
**更新时间**: 2025-11-09
**作者**: Claude Code Enhanced
