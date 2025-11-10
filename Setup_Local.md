# C#程序运行指南 - 针对AI助手

## 项目概述

本项目包含一个用C#编写的Word文档格式分析工具，能够提取和分析`.docx`文件的详细格式信息。

## 项目结构

```
/Users/CodeProjects/OfficeTest/
├── test.docx                           # 待分析的Word文档
├── DocxFormatExtractor/                # C#项目目录 
│   ├── EnhancedProgram.cs              # 增强版程序（当前使用）
│   ├── DocxFormatExtractor.csproj      # 项目配置文件
│   └── bin/Debug/net8.0/               # 编译输出目录
│       └── DocxFormatExtractor         # 可执行文件（已编译）
├── format_output_enhanced.txt          # 文本格式报告（运行后生成）
└── format_output_enhanced.json         # JSON格式报告（运行后生成）
```

## 重要提示：无需安装.NET SDK

**关键信息**：项目已经包含编译好的可执行文件，你**不需要**安装.NET SDK或运行`dotnet run`命令。

## 如何运行程序

### 方法1：直接运行已编译的可执行文件（推荐）

```bash
./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
```

**为什么这个方法有效：**
- 项目已经预编译好了
- 可执行文件位于：`DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor`
- 这是一个独立的可执行文件，可以直接运行

### 方法2：如果系统安装了.NET SDK（可选）

如果用户的系统恰好安装了.NET 8.0 SDK，也可以使用：

```bash
cd DocxFormatExtractor
dotnet run
```

但这**不是必需的**，因为已经有编译好的版本。

## 预期输出

程序运行后会：

1. **在控制台显示进度信息：**
```
开始提取文档格式信息...
使用 Open XML SDK 3.x

1/10 提取文档属性...
2/10 提取样式信息...
3/10 提取段落和文本...
4/10 提取表格...
5/10 提取节信息...
6/10 提取图片...
7/10 提取页眉页脚...
8/10 提取超链接和书签...
9/10 提取字体和编号...
10/10 提取主题和批注...
文本报告已保存到: /Users/CodeProjects/OfficeTest/format_output_enhanced.txt
JSON报告已保存到: /Users/CodeProjects/OfficeTest/format_output_enhanced.json
提取完成！
```

2. **生成两个报告文件：**
   - `format_output_enhanced.txt` - 人类可读的文本格式报告
   - `format_output_enhanced.json` - 机器可读的JSON格式报告

## 查看分析结果

运行程序后，使用Read工具查看生成的报告：

```bash
# 查看文本报告
Read /Users/CodeProjects/OfficeTest/format_output_enhanced.txt

# 查看JSON报告
Read /Users/CodeProjects/OfficeTest/format_output_enhanced.json
```

## 报告内容说明

分析报告包含以下信息：

1. **文档属性** - 创建者、修改时间、页数、字数等
2. **样式信息** - 所有样式定义（标题、正文等）
3. **段落详情** - 每个段落的格式和内容
4. **表格结构** - 表格数量、行数、内容
5. **图片信息** - 图片类型、大小、尺寸
6. **节信息** - 页面大小、边距、方向
7. **超链接** - 文档中的所有链接
8. **书签** - 文档中的书签
9. **页眉页脚** - 页眉页脚内容
10. **批注** - 文档批注（如有）
11. **主题** - 文档主题
12. **字体** - 使用的字体列表
13. **编号** - 列表和编号定义

## 常见错误及解决方案

### 错误1: "command not found: dotnet"

**原因：** 你试图使用`dotnet run`但系统没有安装.NET SDK

**解决方案：** 使用方法1（直接运行已编译的可执行文件）

```bash
./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
```

### 错误2: "Permission denied"

**原因：** 可执行文件没有执行权限

**解决方案：** 添加执行权限

```bash
chmod +x ./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
```

### 错误3: 找不到test.docx

**原因：** 程序在`EnhancedProgram.cs`中硬编码了test.docx的路径

**解决方案：** 确保test.docx文件存在于项目根目录

```bash
ls -la /Users/CodeProjects/OfficeTest/test.docx
```

## 分析不同的Word文档

目前程序默认分析`test.docx`。如果需要分析其他文档：

1. **临时方案**：将要分析的文档重命名为`test.docx`
2. **修改源代码**：编辑`EnhancedProgram.cs`中的文件路径，然后重新编译

## AI助手操作建议

作为AI助手，当用户要求运行C#程序时：

1. **首先检查是否有已编译的可执行文件**
   ```bash
   ls -la DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
   ```

2. **如果存在，直接运行**
   ```bash
   ./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor
   ```

3. **不要假设需要安装.NET SDK** - 这会浪费时间

4. **运行成功后，主动读取生成的报告文件**
   ```bash
   Read /Users/CodeProjects/OfficeTest/format_output_enhanced.txt
   ```

5. **向用户展示关键信息**，如文档统计、样式数量、表格数量等

## 技术细节

- **使用的框架**：.NET 8.0
- **依赖库**：DocumentFormat.OpenXml (Open XML SDK)
- **目标平台**：跨平台（macOS/Linux/Windows）
- **输入格式**：.docx (Office Open XML)
- **输出格式**：TXT + JSON

## 验证程序是否成功运行

程序成功运行的标志：

1. ✓ 控制台输出10个步骤的进度信息
2. ✓ 生成`format_output_enhanced.txt`文件
3. ✓ 生成`format_output_enhanced.json`文件
4. ✓ 文本报告包含完整的文档分析结果（约171行）

## 快速测试流程

```bash
# 1. 运行程序
./DocxFormatExtractor/bin/Debug/net8.0/DocxFormatExtractor

# 2. 确认文件生成
ls -lh format_output_enhanced.*

# 3. 查看报告行数
wc -l format_output_enhanced.txt

# 4. 读取报告内容
head -50 format_output_enhanced.txt
```

## 总结

记住最重要的一点：**项目已经包含编译好的可执行文件，直接运行即可！**

不需要：
- ❌ 安装.NET SDK
- ❌ 运行`dotnet build`
- ❌ 运行`dotnet run`
- ❌ 配置开发环境

只需要：
- ✓ 运行已编译的可执行文件
- ✓ 读取生成的报告文件
- ✓ 向用户展示分析结果

祝你成功！
