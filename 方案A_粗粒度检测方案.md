# 方案A：粗粒度论文格式检测方案

## 方案概述

**设计理念**：脚本只负责数据提取和单位转换，输出简洁的"格式数据"；AI负责所有格式判断、规范对比和完成度计算。



---

# PART A: 给 Claude Code Web 的技术指导

> 本部分用于指导编写 `extract_format_simple.py` 脚本

## 1. 核心任务

处理流程：
1. **【已完成】** 所有版本的格式数据已提取完成，位于 `batch_output/` 目录：
   - `v01_*_format_output.json`
   - `v02_*_format_output.json`
   - ...
   - `v13_*_format_output.json`
2. **【Python】** 读取 `batch_output/` 中的JSON文件，提取格式信息（**只提取原始数据，不做格式检查**）
3. **【Python】** 输出简洁的 `format_data_vXX.json`

输出：
- `format_data_v01.json`
- `format_data_v02.json`
- ...

**目的**：提供给AI足够的格式数据，让AI自己判断是否符合规范。

## 2. 输入数据说明

✅ **格式数据已就绪**：所有版本的文档格式已通过C#程序批量提取完成

- **输入目录**：`batch_output/`
- **文件列表**：
  - `v01_11225c0_deep_learning_thesis_format_output.json`
  - `v02_6d16af1_deep_learning_thesis_format_output.json`
  - ...
  - `v13_b2eb3f6_deep_learning_thesis_format_output.json`

**注**：每个JSON文件约35-45万字节，包含完整的文档格式信息（段落、样式、表格、节、页眉页脚等）。

## 3. Python处理逻辑（处理JSON）

### 3.1 单位换算表

```python
# Twips → cm (1 cm = 567 twips)
cm = twips / 567

# 半磅 → pt → 中文字号
pt = half_point / 2
SIZE_MAP = {16: "三号", 15: "小三", 14: "四号", 12: "小四", 10.5: "五号", 9: "小五"}

# 行距倍数
multiplier = twips / 240

# 段前/段后间距 → 磅
pt = twips / 20

# 首行缩进 → 字符
chars = twips / 240
```

**重点**：只做单位换算，**不判断是否符合规范**。

### 3.2 对齐方式映射

```python
ALIGNMENT_MAP = {
    'JustificationValues { }': '居中',
    'LeftValues { }': '左对齐',
    'RightValues { }': '右对齐',
    'BothValues { }': '两端对齐',
    '': '顶格/默认'
}
```

### 3.3 样式继承解析

```python
def get_effective_font(para, styles):
    """解析最终生效的字体（处理继承）"""
    run = para['Runs'][0] if para['Runs'] else {}

    # 优先级：Run → StyleId → Normal → 默认值
    font_east = run.get('FontNameEastAsia')
    font_ascii = run.get('FontNameAscii')
    font_size = run.get('FontSize')

    if not font_east or not font_size:
        style_id = para.get('StyleId') or 'Normal'
        style = styles.get(style_id, styles.get('Normal', {}))

        if style.get('RunProperties'):
            rp = style['RunProperties']
            font_east = font_east or rp.get('FontNameEastAsia', '宋体')
            font_ascii = font_ascii or rp.get('FontNameAscii', 'Times New Roman')
            font_size = font_size or rp.get('FontSize', '24')

    return {
        'chinese': font_east or '宋体',
        'english': font_ascii or 'Times New Roman',
        'size': int(font_size) if font_size else 24
    }
```

### 3.4 段落分类

```python
def classify_paragraph(para):
    """段落类型分类（同方案B）"""
    text = para['Text'].strip()

    if 'PAGEREF' in text:
        return 'TOC_REF'

    if text in ['摘  要']: return 'ABSTRACT_CN_TITLE'
    if text == 'Abstract': return 'ABSTRACT_EN_TITLE'
    if text in ['目  录']: return 'TOC_TITLE'
    if text == '参考文献': return 'REFERENCE_TITLE'
    if text in ['致  谢']: return 'ACKNOWLEDGEMENT_TITLE'
    if text in ['附  录']: return 'APPENDIX_TITLE'

    if re.match(r'^第\d+章\s+', text): return 'HEADING_1'
    if re.match(r'^\d+\.\d+\s+\S', text): return 'HEADING_2'
    if re.match(r'^\d+\.\d+\.\d+\s+\S', text): return 'HEADING_3'
    if text.startswith('关键词：'): return 'KEYWORDS_CN'
    if text.startswith('Key words'): return 'KEYWORDS_EN'
    if re.match(r'^图\d+-\d+', text): return 'FIGURE_CAPTION'
    if re.match(r'^表\d+-\d+', text): return 'TABLE_CAPTION'
    if re.match(r'^\[\d+\]', text): return 'REFERENCE_ITEM'

    return 'BODY'
```

## 3. 输出JSON结构

**关键区别**：只输出格式数据（value），**不输出expected和match**

```python
{
  "page_setup": {
    "margin_top": "2.0 cm",
    "margin_bottom": "2.0 cm",
    "margin_left": "2.5 cm",
    "margin_right": "2.0 cm",
    "page_width": "21.6 cm",
    "page_height": "28.0 cm",
    "line_spacing_default": "1.5倍"
  },

  "sections": {
    "abstract_cn": {
      "title": {
        "text": "摘  要",
        "font": "宋体",
        "size": "三号(16pt)",
        "bold": true,
        "alignment": "居中",
        "spacing_before": "12磅"
      },
      "content": {
        "count": 2,
        "font": "宋体",
        "size": "小四(12pt)",
        "first_line_indent": "2字符",
        "alignment": "两端对齐"
      },
      "keywords": {
        "text": "关键词：深度学习；图像识别；卷积神经网络；ResNet；计算机视觉",
        "label_bold": true,
        "separator": "；",
        "keyword_count": 5
      }
    },
    "abstract_en": {...},
    "toc": {...},
    "main": {
      "h1": {
        "count": 6,
        "items": [
          {
            "index": 69,
            "text": "第1章 绪论",
            "font": "宋体(继承)",
            "size": "三号(16pt)",
            "bold": true,
            "alignment": "居中",
            "spacing_before": "12磅"
          },
          // ... 其他5个
        ]
      },
      "h2": {...},
      "h3": {...},
      "body": {
        "count": 156,
        "items": [
          {
            "index": 70,
            "text": "图像识别是计算机视觉领域的核心任务...",
            "font": "宋体",
            "size": "小四(12pt)",
            "first_line_indent": "2字符",
            "line_spacing": "1.5倍",
            "alignment": "两端对齐"
          },
          // ... 其他155个
        ]
      }
    },
    "figures": {
      "count": 2,
      "items": [...]
    },
    "tables": {
      "count": 5,
      "items": [...]
    },
    "references": {...},
    "headers_footers": {...},
    "acknowledgement": {...},
    "appendix": {...}
  }
}
```

**设计原则**：
1. 只输出格式数据，不做任何符合性判断
2. 使用易读的单位（cm、磅、中文字号）
3. 输出所有段落的完整格式信息
4. 让AI自己对比规范和实际数据


---

# PART B: 给 Agent AI 的提示词

> 此提示词将直接复制给另一个AI用于计算完成度

````markdown
# 浙江师范大学本科毕业论文格式检测

## 角色
你是浙江师范大学本科毕业论文格式检测专家。

## 任务
对比论文格式规范和实际文档格式数据，**独立判断每个格式项是否符合规范**，计算完成度并生成详细报告。

---

## 输入1：格式规范

/Users/CodeProjects/OfficeTest/浙江师范大学本科毕业论文（设计）格式规范_filtered.md

---

## 输入2：实际文档格式数据

```json
<此处粘贴脚本生成的 format_data_vXX.json>
```

**JSON说明**：
- 脚本已完成：单位转换（cm、磅、中文字号）、段落分类、样式继承解析
- 脚本**未做**：格式符合性判断（没有expected和match字段）
- **你需要做**：逐项对比规范，判断每个格式是否符合要求

---

## 你的工作

### 步骤1：理解格式规范

仔细阅读格式规范MD文件，提取所有可检查的格式要求。例如：
- 页边距：上2.0cm、下2.0cm、左2.5cm、右2.0cm
- 一级标题：宋体、三号、加粗、居中、段前12磅
- 二级标题：宋体、四号、加粗、顶格
- 正文：宋体、小四、首行缩进2字符、1.5倍行距、两端对齐
- ...

### 步骤2：逐项对比检查

对JSON中的每个格式项，与规范对比：

**示例1：页边距检查**
- 规范要求：上边距 2.0 cm
- 实际数据：`"margin_top": "2.0 cm"`
- 判断：✓ 符合

**示例2：一级标题检查**
- 规范要求：宋体、三号、加粗、居中、段前12磅
- 实际数据（第1章）：
  ```json
  {
    "text": "第1章 绪论",
    "font": "宋体(继承)",
    "size": "三号(16pt)",
    "bold": true,
    "alignment": "居中",
    "spacing_before": "12磅"
  }
  ```
- 逐项判断：
  - 字体：宋体 ✓
  - 字号：三号 ✓
  - 加粗：true ✓
  - 对齐：居中 ✓
  - 段前：12磅 ✓
- 结论：完全符合

**示例3：如果有错误**
- 实际数据（某标题）：
  ```json
  {
    "font": "Times New Roman",
    "size": "小三(15pt)",
    "bold": false
  }
  ```
- 判断：
  - 字体：应为宋体，实际为Times New Roman ✗
  - 字号：应为三号，实际为小三 ✗
  - 加粗：应为true，实际为false ✗

### 步骤3：计算完成度

**权重分配**：
- 页面设置: 10%
- 中文摘要: 8%
- 英文摘要: 8%
- 目录: 8%
- 正文格式: 40% (一级标题10% + 二级标题8% + 三级标题7% + 正文段落15%)
- 图片: 8%
- 表格: 8%
- 参考文献: 5%
- 页眉页脚: 3%
- 致谢: 1%
- 附录: 1%

**计算公式**：
```
类别完成度 = (符合规范的项数 / 总检查项数) × 100%
总体完成度 = Σ(类别完成度 × 类别权重)
```

**示例**：
- 一级标题共6个，每个检查5项（字体、字号、加粗、对齐、段前）= 30项
- 如果28项符合，2项不符合
- 一级标题完成度 = 28/30 = 93.3%

### 步骤4：生成报告

严格按以下JSON格式输出：

```json
{
  "overall_completion": 87.5,
  "completion_by_category": {
    "页面设置": 100.0,
    "中文摘要": 100.0,
    "英文摘要": 100.0,
    "目录": 90.0,
    "正文格式": 85.0,
    "图片": 100.0,
    "表格": 100.0,
    "参考文献": 95.0,
    "页眉页脚": 100.0,
    "致谢": 100.0,
    "附录": 100.0
  },
  "failed_rules": [
    {
      "category": "正文-一级标题",
      "rule": "字体应为宋体",
      "actual": "Times New Roman",
      "locations": ["第1章 绪论", "第2章 深度学习与图像识别基础"],
      "severity": "critical",
      "suggestion": "将一级标题字体改为宋体"
    },
    {
      "category": "目录",
      "rule": "目录条目字号应为五号",
      "actual": "小四",
      "severity": "medium",
      "suggestion": "调整目录字号为五号"
    }
  ],
  "summary": {
    "total_issues": 2,
    "critical_issues": 1,
    "medium_issues": 1,
    "low_issues": 0
  },
  "recommendations": [
    "严重问题：一级标题字体错误（2处），需改为宋体",
    "中等问题：目录字号不符合要求",
    "整体评价：文档格式基本符合要求，修正2处问题即可达到95%以上完成度"
  ]
}
```

---

## 注意事项

1. **独立判断**：脚本未提供expected和match，你需要根据规范独立判断每项
2. **全面检查**：检查items数组中的所有段落，不要遗漏
3. **精确对比**：字体、字号、对齐等必须完全匹配规范要求
4. **问题分级**：
   - `critical`（严重）：字体、字号、页边距等核心格式错误
   - `medium`（中等）：间距、对齐等次要格式偏差
   - `low`（轻微）：个别段落的小问题
5. **简洁报告**：recommendations不超过5条

现在开始检测。
````

