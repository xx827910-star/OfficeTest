# V16 格式检测分析文档

## 1. 页面设置检查

### 规范要求：
- 纸张: A4 (21.0cm × 29.7cm)
- 页边距: 上2.0cm, 下2.0cm, 左2.5cm, 右2.0cm
- 文档网格: 无网格
- 全文汉字字体: 宋体
- 章节序号、字母与数字字体: Times New Roman
- 行距: 1.5倍行距
- 对齐: 两端对齐

### 实际数据 (v16):
- page_width: "21.0 cm" ✓
- page_height: "29.7 cm" ✓
- margin_top: "2.0 cm" ✓
- margin_bottom: "2.0 cm" ✓
- margin_left: "2.5 cm" ✓
- margin_right: "2.0 cm" ✓
- defaults.run.font_chinese: "宋体" ✓
- defaults.run.font_english: "Times New Roman" ✓
- defaults.paragraph.line_spacing: "1.5倍" ✓

### 问题：
无

---

## 2. 中文摘要检查

### 规范要求：
- 标题: "摘 要" (宋体, 三号, 加粗, 居中, 段前12磅, 两字间空两格)
- 内容: 宋体, 小四号, 首行缩进2字符, 两端对齐
- 关键词标签: "关键词" (宋体, 小四号, 加粗)
- 关键词内容: 宋体, 小四号, 3-5个, 分号分隔, 首行缩进2字符
- 无页眉
- 页码: 罗马数字 (Times New Roman, 小五, 居中)

### 实际数据 (v16):
#### 标题 (abstract_cn.title):
- text: "摘  要" ✓ (两字间有空格)
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓

#### 内容 (abstract_cn.content):
- count: 2
- font: "宋体" ✓
- size: "小四(12.0pt)" ✓
- first_line_indent: "2字符" ✓
- alignment: "两端对齐" ✓

#### 关键词 (abstract_cn.keywords):
- text: "关键词：深度学习；图像识别；卷积神经网络；ResNet；计算机视觉"
- separator: "；" ✓
- keyword_count: 5 ✓
- label_bold: true ✓
- first_line_indent: "2字符" ✓

### 问题：
无

---

## 3. 英文摘要检查

### 规范要求：
- 标题: "Abstract" (Times New Roman, 三号, 加粗, 居中, 段前12磅)
- 内容: Times New Roman, 小四号, 首行缩进2字符, 两端对齐
- 关键词标签: "Key words" (Times New Roman, 小四号)
- 关键词内容: Times New Roman, 小四号, 分号分隔, 首行缩进2字符
- 无页眉
- 页码: 罗马数字 (Times New Roman, 小五, 居中, 延续中文摘要)

### 实际数据 (v16):
#### 标题 (abstract_en.title):
- text: "Abstract" ✓
- font: "Times New Roman" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓

#### 内容 (abstract_en.content):
- count: 2
- font: "Times New Roman" ✓
- size: "小四(12.0pt)" ✓
- first_line_indent: "2字符" ✓
- alignment: "两端对齐" ✓

#### 关键词 (abstract_en.keywords):
- text: "Key words: Deep Learning; Image Recognition; Convolutional Neural Networks; ResNet; Computer Vision"
- separator: ";" ✓
- keyword_count: 5 ✓
- label_bold: true ✓ (需要验证 - 但数据没有显示 Key words 部分未加粗的问题)
- first_line_indent: "2字符" ✓

### 问题：
无

---

## 4. 目录检查

### 规范要求：
- 标题: "目 录" (宋体, 三号, 加粗, 居中, 段前12磅, 两字间空两格)
- 一级标题: 宋体, 五号, 顶格
- 二级标题: 宋体, 五号, 缩进1字符
- 三级标题: 宋体, 五号, 缩进2字符
- 无页眉
- 页码: 罗马数字 (Times New Roman, 小五, 居中, 从i开始)

### 实际数据 (v16):
#### 标题 (toc.title):
- text: "目  录" ✓
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓

#### 目录条目:
- items_count: 52

##### Profile 1 (一级标题):
- toc_level: 1
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- left_indent: "" ✓ (顶格)
- count: 9

##### Profile 2 (二级标题):
- toc_level: 2
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- left_indent: "1.0字符" ✓
- count: 21

##### Profile 3 (三级标题):
- toc_level: 3
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- left_indent: "2.0字符" ✓
- count: 22

### 问题：
无

---

## 5. 正文格式检查

### 规范要求：
- 一级标题: 宋体, 三号, 加粗, 居中, 段前12磅, 章序号与章名间空一个字符
- 二级标题: 宋体, 四号, 加粗, 顶格, 序号和名称空一个字符
- 三级标题: 宋体, 小四号, 加粗, 顶格, 序号和名称空一个字符
- 正文: 宋体, 小四号, 首行缩进2字符, 1.5倍行距, 两端对齐
- 页眉: 论文题目 (宋体, 小五, 居中)
- 页码: 阿拉伯数字 (Times New Roman, 小五, 居中, 从1开始)

### 实际数据 (v16):
#### 一级标题 (main.h1):
- count: 6
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓
- sample_text: "第1章 绪论" (章序号与章名间有空格) ✓

#### 二级标题 (main.h2):
- count: 21
- font: "宋体" ✓
- size: "四号(14.0pt)" ✓
- bold: true ✓
- alignment: "左对齐" ✓ (顶格)
- sample_text: "1.1 研究背景与意义" (序号和名称间有空格) ✓

#### 三级标题 (main.h3):
- count: 22
- font: "宋体" ✓
- size: "小四(12.0pt)" ✓
- bold: true ✓
- alignment: "左对齐" ✓ (顶格)
- sample_text: "1.1.1 理论意义" (序号和名称间有空格) ✓

#### 正文段落 (main.body):
##### Profile 1 (主要正文):
- count: 65
- font: "宋体" ✓
- size: "小四(12.0pt)" ✓
- first_line_indent: "2字符" ✓
- line_spacing: "1.5倍" ✓
- alignment: "两端对齐" ✓

##### Profile 2 (公式段落 - deviations):
- count: 10
- font: "宋体" ✗ (应为 Times New Roman 五号)
- size: "五号(10.5pt)" ✓
- first_line_indent: "" ✓ (公式居中)
- alignment: "未提供"

### 问题：
1. **公式段落字体问题**: deviations 中的 body_profile_2 (公式段落) 字体为"宋体"，规范要求公式字体应为 Times New Roman 五号。(10个公式段落，索引: 101, 105, 109, 115, 119, 127, 131, 136, 164, 174)

---

## 6. 图片检查

### 规范要求：
- 编号和图名: 五号宋体
- 图号和图名: 居图下方正中
- 按章编号 (如 图1-1)
- 资料来源: 小五宋体, 置图左下方
- 图与上下正文间各空一行

### 实际数据 (v16):
- count: 2

#### 图1-1 (index 88):
- text: "图1-1 本文研究技术路线图"
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- alignment: "居中" ✓
- blank_before: true ✓
- blank_after: false ✗ (应为 true)
- source.text: "来源：自制"
- source.font: "宋体" ✓
- source.size: "小五(9.0pt)" ✓

#### 图3-1 (index 154):
- text: "图3-1 VGG-16网络架构图"
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- alignment: "居中" ✓
- blank_before: true ✓
- blank_after: false ✗ (应为 true)
- source.text: "来源：Simonyan et al., 2014"
- source.font: "宋体" ✓
- source.size: "小五(9.0pt)" ✓

### 问题：
根据规范第69行"④图与上下正文之间各空一行"，但 prompt.md 第 30 行特别说明："图片/表格空行: 脚本无法准确识别与正文之间的空行, 若仅因空行与规范不匹配, 请记 0 权重。"

因此，blank_after 的问题记为 0 权重，不记录为格式错误。

---

## 7. 表格检查

### 规范要求：
- 表号和表名: 五号宋体, 居表上方正中, 中间空一格
- 按章编号 (如 表1-1)
- 三线表: 上下线1.5磅, 中间线0.5磅, 无竖线
- 表中字体: 五号宋体
- 资料来源: 小五宋体, 置表格左下方
- 表与上下正文间各空一行

### 实际数据 (v16):
- count: 5

#### 表格标题格式 (defaults.caption):
- font: "宋体" ✓
- size: "五号(10.5pt)" ✓
- bold: false ✓
- alignment: "居中" ✓
- blank_before: true ✓
- blank_after: false ✗ (但根据特别说明记0权重)

#### 资料来源格式 (defaults.source):
- font: "宋体" ✓
- size: "小五(9.0pt)" ✓

#### 表格结构 (structure):
所有5个表格的结构:
- top_border.size: "1.50pt" ✓
- bottom_border.size: "1.50pt" ✓
- inside_horizontal.size: "0.50pt" ✓
- has_inside_vertical: false ✓ (无竖线)
- has_vertical_outer: false ✓

#### 各表格条目:
所有表格的 caption_diff 和 source.diff 都为空 {}, 说明与 defaults 一致 ✓
stats.with_source: 5, without_source: 0 ✓

### 问题：
无 (blank_after 问题记0权重)

---

## 8. 公式检查

### 规范要求：
- 编号: 阿拉伯数字连续编序号, 按章编号 (如 2-1)
- 序号: 加圆括号, 标记于公式后面, 右对齐
- 公式: 居中
- 公式字体: Times New Roman 五号

### 实际数据 (v16):
- count: 10

检查每个公式:
1. paragraph_index: 101, numbering_text: "(2-1)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
2. paragraph_index: 105, numbering_text: "(2-2)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
3. paragraph_index: 109, numbering_text: "(2-3)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
4. paragraph_index: 115, numbering_text: "(2-4)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
5. paragraph_index: 119, numbering_text: "(2-5)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
6. paragraph_index: 127, numbering_text: "(2-6)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
7. paragraph_index: 131, numbering_text: "(2-7)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
8. paragraph_index: 136, numbering_text: "(2-8)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
9. paragraph_index: 164, numbering_text: "(3-1)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓
10. paragraph_index: 174, numbering_text: "(3-2)" ✓, numbering_font: "Times New Roman" ✓, numbering_font_size: "五号(10.5pt)" ✓

注意: equation_font 和 equation_font_size 均为空 ""

### 问题：
**公式字体信息缺失**: 所有10个公式的 equation_font 和 equation_font_size 字段都为空，无法验证公式本身是否使用 Times New Roman 五号字体。但根据正文 body_profile_2 的数据，这些公式段落的字体为"宋体"而非"Times New Roman"，这是一个问题。

关联到正文检查第5项，公式段落应该使用 Times New Roman 五号字体。

---

## 9. 参考文献检查

### 规范要求：
- 标题: "参考文献" (宋体, 三号, 加粗, 居中)
- 内容: 小四宋体
- 格式: 序号加方括号, 左顶格
- 作者3人以上只列前3人, 后加"等"或"et al"
- 每条目以实心点结束

### 实际数据 (v16):
#### 标题 (references.title):
- text: "参考文献" ✓
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓

#### 条目:
- items_count: 8

检查各条目 (items):
1. [1] - font: "宋体" ✓, size: "小四(12.0pt)" ✓, hanging_indent: "" (无悬挂缩进, 规范未要求 ✓), starts_with_bracket: true ✓, sequence_number: 1 ✓, ends_with_period: true ✓
2. [2] - 同上 ✓
3. [3] - 同上 ✓
4. [4] - 同上 ✓
5. [5] - 同上 ✓
6-8: 数据只显示了5条, 但 items_count 为 8

### 问题：
**参考文献数据不完整**: items_count 显示为8条, 但实际 items 数组只有5条数据 (indexes: 245-249)。缺少3条参考文献的详细信息。这可能是数据采样的问题，但无法验证第6-8条是否符合规范。

---

## 10. 页眉页脚检查

### 规范要求：
- 中文摘要: 无页眉, 罗马数字页码 (Times New Roman, 小五, 居中)
- 英文摘要: 无页眉, 罗马数字页码 (Times New Roman, 小五, 居中, 延续)
- 目录: 无页眉, 罗马数字页码 (Times New Roman, 小五, 居中, 从i开始)
- 正文: 页眉为论文题目 (宋体, 小五, 居中), 阿拉伯数字页码 (Times New Roman, 小五, 居中, 从1开始)
- 参考文献: 页眉为"参考文献" (根据特别说明)
- 致谢: 页眉为"致谢"
- 附录: 页眉为"附录"

### 实际数据 (v16):
#### 页眉 (headers):
- header_count: 4
  - [0] text: "基于深度学习的图像识别技术研究与应用", font: "宋体", size: "小五(9.0pt)" ✓, alignment: "居中" ✓
  - [1] text: "参考文献", font: "宋体", size: "小五(9.0pt)" ✓, alignment: "居中" ✓
  - [2] text: "致  谢", font: "宋体", size: "小五(9.0pt)" ✓, alignment: "居中" ✓
  - [3] text: "附  录", font: "宋体", size: "小五(9.0pt)" ✓, alignment: "居中" ✓

#### 页脚 (footers):
- footer_count: 3
  - [0] text: "PAGE \\* roman", font: "Times New Roman", size: "小五(9.0pt)" ✓, alignment: "居中" ✓ (罗马数字)
  - [1] text: "PAGE \\* roman", font: "Times New Roman", size: "小五(9.0pt)" ✓, alignment: "居中" ✓ (罗马数字)
  - [2] text: "PAGE", font: "Times New Roman", size: "小五(9.0pt)" ✓, alignment: "居中" ✓ (阿拉伯数字)

#### 节设置 (section_settings):
只有1个节的信息 (index: 0), 缺少其他节的详细信息，无法确认前置部分是否确实无页眉。

### 问题：
**节设置信息不完整**: section_settings 只有1个条目，无法验证中文摘要、英文摘要、目录等前置部分是否真的无页眉。但从 headers 数据来看，有4个页眉设置，符合规范要求（正文、参考文献、致谢、附录各有页眉）。

---

## 11. 致谢检查

### 规范要求：
- 标题: "致 谢" (宋体, 三号, 居中, 段前12磅, 两字间空两格)
- 内容: 宋体小四号

### 实际数据 (v16):
#### 标题 (acknowledgement.title):
- text: "致  谢" ✓ (两字间有空格)
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓

#### 内容:
- content_count: 3
- content_samples 显示3个段落:
  - font: "宋体" ✓
  - font_english: "Times New Roman" ✗ (英文字体不应该被单独标注为 Times New Roman, 但这可能是英文部分的字体设置，需要核实)
  - size: "小四(12.0pt)" ✓
  - alignment: "两端对齐" ✓
  - first_line_indent: "2字符" ✓

### 问题：
无明显问题 (font_english 为 Times New Roman 符合规范第14行"字母与数字的字体为Times New Roman")

---

## 12. 附录检查

### 规范要求：
- 标题: "附录" (宋体, 三号, 居中, 段前12磅, 两字间空两格)
- 内容: 宋体小四号

### 实际数据 (v16):
#### 标题 (appendix.title):
- text: "附  录" ✓ (两字间有空格)
- font: "宋体" ✓
- size: "三号(16.0pt)" ✓
- bold: true ✓
- alignment: "居中" ✓
- spacing_before: "12磅" ✓

#### 内容:
- content_count: 3
- content_samples 显示3个段落:
  - font: "宋体" ✓
  - font_english: "Times New Roman" ✓ (符合规范)
  - size: "小四(12.0pt)" ✓
  - alignment: "两端对齐" ✓
  - first_line_indent: "2字符" ✓

### 问题：
无

---

## 汇总

### 发现的问题：

1. **公式字体问题** (严重 - critical):
   - 位置: 正文 body_profile_2 (deviations)
   - 规范要求: 公式字体用 Times New Roman 五号
   - 实际: 宋体, 五号(10.5pt)
   - 影响范围: 10个公式段落 (indexes: 101, 105, 109, 115, 119, 127, 131, 136, 164, 174)

2. **参考文献数据不完整** (无法判断):
   - items_count 显示8条，但只提供了5条详细数据
   - 缺少第6-8条的验证

### 检查项统计：

- 页面设置: 检查9项, 问题0个
- 中文摘要: 检查8项, 问题0个
- 英文摘要: 检查8项, 问题0个
- 目录: 检查10项, 问题0个
- 正文格式: 检查20项, 问题1个 (公式字体)
- 图片: 检查12项, 问题0个 (空行问题记0权重)
- 表格: 检查15项, 问题0个 (空行问题记0权重)
- 公式: 检查4项 (编号、编号字体、编号字号、公式字体), 问题1个 (公式字体)
- 参考文献: 检查7项, 问题0个 (数据不完整但可见部分符合)
- 页眉页脚: 检查8项, 问题0个
- 致谢: 检查6项, 问题0个
- 附录: 检查6项, 问题0个

**总计: 检查113项, 发现1个问题**
