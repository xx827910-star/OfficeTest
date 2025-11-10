# 方案B：中粒度论文格式检测方案

## 方案概述

**设计理念**：脚本负责数据提取、单位转换、段落分类和基础格式检查，输出结构化的"格式摘要"；AI专注于对比规范和实际格式，给出完成度评分。


---

# PART A: 给 Claude Code Web 的技术指导

> 本部分用于指导编写 `extract_format.py` 脚本

## 1. 核心任务

处理流程：
1. **【已完成】** 所有版本的格式数据已提取完成，位于 `batch_output/` 目录：
   - `v01_*_format_output.json`
   - `v02_*_format_output.json`
   - ...
   - `v13_*_format_output.json`
2. **【Python】** 读取 `batch_output/` 中的JSON文件，并提取格式信息，执行单位转换、段落分类和基础格式检查
3. **【Python】** 输出 `format_data_vXX.json` 供Agent AI检测

输出：
- `format_data_v01.json`
- `format_data_v02.json`
- ...


**目的**：模拟格式完善的迭代过程，让Agent AI返回每个版本的完成度，观察完成度如何从v01的低水平逐步提升到test的高水平。

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
chars = twips / 240  # 480 twips ≈ 2字符
```

**精度判断**：换算结果是整数或一位小数 → 单位正确，否则用错单位。

### 3.2 对齐方式映射

```python
ALIGNMENT_MAP = {
    'JustificationValues { }': '居中',  # Center
    'LeftValues { }': '左对齐',
    'RightValues { }': '右对齐',
    'BothValues { }': '两端对齐',      # Justify
    '': '顶格/默认'
}
```

**注**：实测 `JustificationValues { }` 在不同上下文可能表示"居中"或"两端对齐"，需结合段落类型判断：
- 标题类段落 → 居中
- 正文类段落 → 两端对齐

### 3.3 样式继承解析

```python
def get_effective_font(para, styles):
    """解析最终生效的字体（处理继承）"""
    run = para['Runs'][0] if para['Runs'] else {}

    # 优先级：Run直接格式 > StyleId样式 > Normal样式 > 默认值
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

    # 边界情况：全空则使用默认值
    return {
        'chinese': font_east or '宋体',
        'english': font_ascii or 'Times New Roman',
        'size': int(font_size) if font_size else 24
    }
```

### 3.4 段落分类

```python
def classify_paragraph(para):
    """段落类型分类"""
    text = para['Text'].strip()

    # 排除目录引用
    if 'PAGEREF' in text:
        return 'TOC_REF'

    # 特殊标题（精确匹配）
    if text in ['摘  要']: return 'ABSTRACT_CN_TITLE'
    if text == 'Abstract': return 'ABSTRACT_EN_TITLE'
    if text in ['目  录']: return 'TOC_TITLE'
    if text == '参考文献': return 'REFERENCE_TITLE'
    if text in ['致  谢']: return 'ACKNOWLEDGEMENT_TITLE'
    if text in ['附  录']: return 'APPENDIX_TITLE'

    # 正则匹配
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

### 3.5 三线表检测

```python
def is_three_line_table(table):
    """判断是否为三线表（无竖线，仅上中下三条横线）"""
    # 方法1：检查HasBorders（粗略判断）
    if not table.get('HasBorders'):
        return False

    # 方法2：检查具体边框设置（需解析XML或额外属性）
    # 三线表特征：
    # - 上边框：1.5磅
    # - 表头下边框：0.5磅
    # - 下边框：1.5磅
    # - 无左右边框、无内部竖线

    # 简化判断：如果HasBorders=True且表格规整，暂定为三线表
    # 让AI在提示词中进一步判断
    return True
```

### 3.6 格式检查与Match判断

```python
def check_format_item(actual, expected):
    """检查单个格式项，返回 {value, expected, match}"""

    # 字符串精确匹配
    if isinstance(expected, str):
        match = (actual == expected)

    # 数值容差匹配（单位正确时误差应为0）
    elif isinstance(expected, (int, float)):
        match = abs(actual - expected) < 0.1

    # 布尔值匹配
    elif isinstance(expected, bool):
        match = (actual == expected)

    return {
        'value': actual,
        'expected': expected,
        'match': match
    }
```

## 3. 输出JSON结构

```python
{
  "page_setup": {
    "margins": {
      "top": {"value": "2.0", "expected": "2.0", "match": true},
      "bottom": {"value": "2.0", "expected": "2.0", "match": true},
      "left": {"value": "2.5", "expected": "2.5", "match": true},
      "right": {"value": "2.0", "expected": "2.0", "match": true}
    },
    "line_spacing": {"value": "1.5倍", "expected": "1.5倍", "match": true}
  },

  "sections": {
    "abstract_cn": {
      "title": {
        "text": "摘  要",
        "format": {
          "font": {"value": "宋体", "expected": "宋体", "match": true},
          "size": {"value": "三号(16pt)", "expected": "三号", "match": true},
          "bold": {"value": true, "expected": true, "match": true},
          "alignment": {"value": "居中", "expected": "居中", "match": true}
        }
      },
      "content": {"count": 2, "format": {...}},
      "keywords": {"text": "...", "format": {...}}
    },
    "abstract_en": {...},
    "toc": {...},
    "main": {
      "h1": {
        "count": 6,
        "items": [
          {"index": 69, "text": "第1章 绪论", "format": {...}},
          {"index": 93, "text": "第2章 深度学习与图像识别基础", "format": {...}},
          {"index": 118, "text": "第3章 主流CNN架构分析", "format": {...}},
          {"index": 155, "text": "第4章 实验设计与结果分析", "format": {...}},
          {"index": 184, "text": "第5章 应用案例研究", "format": {...}},
          {"index": 200, "text": "第6章 总结与展望", "format": {...}}
        ]
      },
      "h2": {
        "count": 21,
        "items": [
          // 全部21个二级标题的详细格式
        ]
      },
      "h3": {
        "count": 22,
        "items": [
          // 全部22个三级标题的详细格式
        ]
      },
      "body": {
        "count": 156,
        "items": [
          // 全部156个正文段落的详细格式
        ]
      }
    },
    "figures": {
      "count": 2,
      "items": [
        // 全部2个图题的详细格式
      ]
    },
    "tables": {
      "count": 5,
      "items": [
        // 全部5个表格的详细格式
      ]
    },
    "references": {
      "count": 28,
      "items": [
        // 全部28个参考文献的详细格式
      ]
    },
    "headers_footers": {...},
    "acknowledgement": {...},
    "appendix": {...}
  },

  "stats": {
    "total_checks": 156,
    "passed": 148,
    "failed": 8,
    "by_category": {...}
  }
}
```

**字段说明**：
- `value`: 实际值（从文档中提取）
- `expected`: 期望值（从规范中读取）
- `match`: 是否匹配（脚本判断）
- `format`: 格式检查结果，包含多个 {value, expected, match} 三元组
- `count`: 该类段落的总数量
- `index`: 段落在原JSON中的索引位置
- `items`: **该类段落的所有项目（不做筛选，全部输出）**

**设计原则**：
1. 使用完整字段名（value/expected/match），便于理解
2. 每个检查项都有 `{value, expected, match}` 三元组
3. **输出所有段落的完整格式信息**（不做样本筛选）
4. 让Agent AI能看到每个段落的详细格式，精确判断完成度

## 4. 实现要点

### 4.1 遍历策略

```python
# 1. 先分类所有段落
classified = classify_all_paragraphs(paragraphs)

# 2. 按类别提取格式
for category, paras in classified.items():
    extract_format_by_category(category, paras, styles)
```

---

# PART B: 给 Agent AI 的提示词

> 此提示词将直接复制给另一个AI用于计算完成度

````markdown
# 浙江师范大学本科毕业论文格式检测

## 角色
你是浙江师范大学本科毕业论文格式检测专家。

## 任务
对比论文格式规范和实际文档格式，计算完成度并生成详细报告。

---

## 输入1：格式规范

OfficeTest/浙江师范大学本科毕业论文（设计）格式规范_filtered.md

---

## 输入2：实际文档格式

```json
<此处粘贴脚本生成的 format_summary_vXX.json>
```

**JSON字段说明**：
- `value`: 实际值（从文档中提取的格式）
- `expected`: 期望值（从格式规范中读取的标准）
- `match`: 是否匹配（脚本已判断，true表示符合，false表示不符合）
- `format`: 格式检查结果，包含多个 {value, expected, match} 三元组
- `count`: 该类段落的总数量（如一级标题共6个）
- `index`: 段落在原始文档中的索引位置
- `items`: **该类段落的所有项目（全部输出，不做筛选）**

**脚本已完成的工作**：
1. 单位换算（twips→cm，半磅→pt→中文字号）
2. 样式继承解析（Run → StyleId → Normal → 默认值）
3. 段落分类（一级标题、二级标题、正文等14种类型）
4. 基础格式检查（所有match字段已由脚本判断）
5. **完整输出（输出每类段落的所有项目，不做样本筛选）**

**注意**：
- 脚本已经做了大量工作，你只需要验证match结果是否合理
- `items`数组包含该类段落的所有项目（如h1有6个，items就包含全部6个的详细格式）
- 每个item都有完整的格式检查结果，你可以精确判断每个段落是否符合规范

---

## 你的工作

### 步骤1：理解JSON结构

理解字段含义：
```json
{"value": "2.0", "expected": "2.0", "match": true}
```
表示：实际值2.0，期望值2.0，匹配成功。

**理解"items"完整输出**：
- 脚本对每类段落（如一级标题共6个）输出全部6个的详细格式
- `count`字段告诉你总数，`items`数组包含所有项目的详细格式
- 每个item都有完整的 {value, expected, match} 格式检查结果
- 你可以逐个检查每个段落是否符合规范，计算该类段落的完成度

### 步骤2：验证脚本判断

检查所有 `match` 字段，判断脚本的匹配是否合理。如果发现明显错误，在报告中指出。

### 步骤3：补充语义判断

脚本无法处理的主观要求（你需要判断）：
- 标题是否"简明扼要，无标点符号"（检查items中的标题文本）
- 关键词数量是否3-5个（检查关键词文本）
- 章序号与章名间距是否1字符（如"第1章 绪论"）

### 步骤4：计算完成度

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
类别完成度 = (该类别通过的检查项 / 该类别总检查项) × 100%
总体完成度 = Σ(类别完成度 × 类别权重)
```

**特殊情况**：
- 如果某个类别不存在（如无附录），该类别按100%计算

### 步骤5：生成报告

严格按以下JSON格式输出：

```json
{
  "overall_completion": 94.5,
  "completion_by_category": {
    "页面设置": 83.3,
    "中文摘要": 100.0,
    "英文摘要": 100.0,
    "目录": 100.0,
    "正文格式": 95.0,
    "图片": 100.0,
    "表格": 100.0,
    "参考文献": 100.0,
    "页眉页脚": 100.0,
    "致谢": 80.0,
    "附录": 100.0
  },
  "failed_rules": [
    {
      "category": "页面设置",
      "rule": "页面尺寸应为A4 (21.0cm × 29.7cm)",
      "actual": "21.6cm × 28.0cm",
      "severity": "medium",
      "suggestion": "调整页面尺寸为标准A4"
    },
    {
      "category": "正文-正文段落",
      "rule": "每段落首行缩进2字符",
      "actual": "2段缺少首行缩进",
      "locations": "索引未提供",
      "severity": "low",
      "suggestion": "为所有正文段落添加首行缩进2字符"
    }
  ],
  "summary": {
    "total_issues": 2,
    "critical_issues": 0,
    "medium_issues": 1,
    "low_issues": 1
  },
  "recommendations": [
    "主要问题：页面尺寸设置不正确，需调整为标准A4",
    "次要问题：2个正文段落缺少首行缩进",
    "整体评价：文档格式基本符合要求，仅需修正2处问题即可达到95%以上完成度"
  ]
}
```

---

## 注意事项

1. **严格对照规范**：每个判断都要有规范依据，引用具体条款
2. **信任脚本判断**：脚本已做单位转换和精确匹配，除非明显错误，否则采纳其判断
3. **完成度精确计算**：使用给定权重，精确到小数点后一位
4. **问题分级**：
   - `critical`（严重）：页边距、字体、字号严重错误
   - `medium`（中等）：页面尺寸、间距偏差
   - `low`（轻微）：个别段落格式问题
5. **简洁明了**：recommendations不超过5条，每条一句话

现在开始检测。
````

