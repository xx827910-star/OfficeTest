# 方案C：细粒度论文格式检测方案

## 方案概述

**设计理念**：脚本负责几乎所有工作（数据提取、格式检查、规则匹配、完成度计算），输出详细的"检查清单"；AI只需验证脚本结果并生成报告。



---

# PART A: 给 Claude Code Web 的技术指导

> 本部分用于指导编写 `extract_format_detailed.py` 脚本

## 1. 核心任务

处理流程：
1. **【已完成】** 所有版本的格式数据已提取完成，位于 `batch_output/` 目录：
   - `v01_*_format_output.json`
   - `v02_*_format_output.json`
   - ...
   - `v13_*_format_output.json`
2. **【Python】** 读取 `batch_output/` 中的JSON文件，并提取格式信息
3. **【Python】** 对照内置规则库，逐条检查格式
4. **【Python】** 计算每个类别的完成度
5. **【Python】** 输出详细的 `format_data_vXX.json`

输出：
- `format_data_v01.json`
- `format_data_v02.json`
- ...

**目的**：提供给AI完整的检查结果，AI只需验证和生成报告。

## 2. 输入数据说明

✅ **格式数据已就绪**：所有版本的文档格式已通过C#程序批量提取完成

- **输入目录**：`batch_output/`
- **文件列表**：
  - `v01_11225c0_deep_learning_thesis_format_output.json`
  - `v02_6d16af1_deep_learning_thesis_format_output.json`
  - ...
  - `v13_b2eb3f6_deep_learning_thesis_format_output.json`

**注**：每个JSON文件约35-45万字节，包含完整的文档格式信息（段落、样式、表格、节、页眉页脚等）。

## 3. Python处理逻辑（规则引擎）

### 3.1 规则库设计 (参照浙江师范大学本科毕业论文（设计）格式规范_filtered.md设计)

```python
# 内置格式规则库
FORMAT_RULES = {
    "page_setup": [
        {
            "id": "R001",
            "name": "页边距-上",
            "description": "上边距应为 2.0 cm",
            "check": lambda data: abs(float(data['margin_top']) - 2.0) < 0.1,
            "get_actual": lambda data: data['margin_top'],
            "expected": "2.0 cm"
        },
        {
            "id": "R002",
            "name": "页边距-下",
            "description": "下边距应为 2.0 cm",
            "check": lambda data: abs(float(data['margin_bottom']) - 2.0) < 0.1,
            "get_actual": lambda data: data['margin_bottom'],
            "expected": "2.0 cm"
        },
        # ... 更多页面设置规则
    ],

    "abstract_cn_title": [
        {
            "id": "R010",
            "name": "摘要标题-字体",
            "description": "摘要标题字体应为宋体",
            "check": lambda data: data['font'] == '宋体',
            "get_actual": lambda data: data['font'],
            "expected": "宋体"
        },
        {
            "id": "R011",
            "name": "摘要标题-字号",
            "description": "摘要标题字号应为三号",
            "check": lambda data: '三号' in data['size'],
            "get_actual": lambda data: data['size'],
            "expected": "三号"
        },
        {
            "id": "R012",
            "name": "摘要标题-加粗",
            "description": "摘要标题应加粗",
            "check": lambda data: data['bold'] == True,
            "get_actual": lambda data: data['bold'],
            "expected": True
        },
        # ... 更多摘要标题规则
    ],

    "heading_1": [
        {
            "id": "R050",
            "name": "一级标题-字体",
            "description": "一级标题字体应为宋体",
            "check": lambda item: '宋体' in item['font'],
            "get_actual": lambda item: item['font'],
            "expected": "宋体",
            "check_all": True  # 需要检查所有一级标题
        },
        {
            "id": "R051",
            "name": "一级标题-字号",
            "description": "一级标题字号应为三号",
            "check": lambda item: '三号' in item['size'],
            "get_actual": lambda item: item['size'],
            "expected": "三号",
            "check_all": True
        },
        # ... 更多一级标题规则
    ],

    # ... 更多类别的规则
}
```

### 3.2 规则检查引擎

```python
def check_rule(rule, data):
    """检查单条规则"""
    try:
        passed = rule['check'](data)
        actual = rule['get_actual'](data)
        return {
            'rule_id': rule['id'],
            'rule_name': rule['name'],
            'description': rule['description'],
            'expected': rule['expected'],
            'actual': actual,
            'passed': passed
        }
    except Exception as e:
        return {
            'rule_id': rule['id'],
            'rule_name': rule['name'],
            'description': rule['description'],
            'expected': rule['expected'],
            'actual': f"检查失败: {str(e)}",
            'passed': False
        }

def check_category_rules(category_rules, data, items=None):
    """检查某个类别的所有规则"""
    results = []

    for rule in category_rules:
        if rule.get('check_all') and items:
            # 需要检查所有项目（如所有一级标题）
            for idx, item in enumerate(items):
                result = check_rule(rule, item)
                result['item_index'] = item.get('index', idx)
                result['item_text'] = item.get('text', '')[:50]
                results.append(result)
        else:
            # 只检查单个数据（如页面设置）
            result = check_rule(rule, data)
            results.append(result)

    return results
```

### 3.3 完成度计算

```python
def calculate_completion(check_results):
    """计算完成度"""
    total = len(check_results)
    passed = sum(1 for r in check_results if r['passed'])

    return {
        'total': total,
        'passed': passed,
        'failed': total - passed,
        'completion': round(passed / total * 100, 1) if total > 0 else 100.0
    }

def calculate_overall_completion(category_completions, weights):
    """计算总体完成度"""
    overall = 0.0
    for category, weight in weights.items():
        completion = category_completions.get(category, {}).get('completion', 100.0)
        overall += completion * weight / 100

    return round(overall, 1)
```

## 3. 输出JSON结构

**关键特点**：包含完整的规则检查清单和预计算的完成度

```python
{
  "document_info": {
    "version": "v01",
    "total_paragraphs": 180,
    "check_timestamp": "2025-01-10 12:00:00"
  },

  "rule_checks": {
    "page_setup": {
      "rules": [
        {
          "rule_id": "R001",
          "rule_name": "页边距-上",
          "description": "上边距应为 2.0 cm",
          "expected": "2.0 cm",
          "actual": "2.0 cm",
          "passed": true
        },
        {
          "rule_id": "R002",
          "rule_name": "页边距-下",
          "description": "下边距应为 2.0 cm",
          "expected": "2.0 cm",
          "actual": "1.8 cm",
          "passed": false
        },
        // ... 其他页面设置规则
      ],
      "completion": {
        "total": 6,
        "passed": 5,
        "failed": 1,
        "completion": 83.3
      }
    },

    "abstract_cn": {
      "title": {
        "rules": [
          {
            "rule_id": "R010",
            "rule_name": "摘要标题-字体",
            "expected": "宋体",
            "actual": "宋体",
            "passed": true
          },
          // ... 其他标题规则
        ],
        "completion": {"total": 6, "passed": 6, "failed": 0, "completion": 100.0}
      },
      "content": {...},
      "keywords": {...}
    },

    "main_content": {
      "h1": {
        "rules": [
          {
            "rule_id": "R050",
            "rule_name": "一级标题-字体",
            "expected": "宋体",
            "actual": "Times New Roman",
            "passed": false,
            "item_index": 69,
            "item_text": "第1章 绪论"
          },
          {
            "rule_id": "R050",
            "rule_name": "一级标题-字体",
            "expected": "宋体",
            "actual": "Times New Roman",
            "passed": false,
            "item_index": 93,
            "item_text": "第2章 深度学习与图像识别基础"
          },
          // ... 所有一级标题的所有规则检查（6个标题 × 5条规则 = 30项）
        ],
        "completion": {"total": 30, "passed": 24, "failed": 6, "completion": 80.0}
      },
      "h2": {...},
      "h3": {...},
      "body": {...}
    },

    "figures": {...},
    "tables": {...},
    "references": {...},
    "headers_footers": {...},
    "acknowledgement": {...},
    "appendix": {...}
  },

  "completion_summary": {
    "by_category": {
      "页面设置": 83.3,
      "中文摘要": 100.0,
      "英文摘要": 100.0,
      "目录": 90.0,
      "正文格式": 82.5,
      "图片": 100.0,
      "表格": 95.0,
      "参考文献": 100.0,
      "页眉页脚": 100.0,
      "致谢": 100.0,
      "附录": 100.0
    },
    "overall": 91.2,
    "total_rules": 245,
    "passed_rules": 223,
    "failed_rules": 22
  },

  "failed_rules_detail": [
    {
      "rule_id": "R002",
      "category": "页面设置",
      "rule_name": "页边距-下",
      "expected": "2.0 cm",
      "actual": "1.8 cm",
      "severity": "medium"
    },
    {
      "rule_id": "R050",
      "category": "正文-一级标题",
      "rule_name": "一级标题-字体",
      "expected": "宋体",
      "actual": "Times New Roman",
      "locations": [
        {"index": 69, "text": "第1章 绪论"},
        {"index": 93, "text": "第2章 深度学习与图像识别基础"},
        // ... 所有失败的位置
      ],
      "severity": "critical"
    }
  ]
}
```

**设计原则**：
1. 完整的规则检查清单（每条规则都有检查结果）
2. 预计算的完成度（按类别）
3. 详细的失败规则信息（包含位置）
4. AI只需验证结果，不需要重新计算

## 4. 实现要点

### 4.1 规则库完整性

```python
# 确保规则库覆盖所有格式要求
# 规则数量估计：
# - 页面设置: 6条
# - 中文摘要: 12条
# - 英文摘要: 10条
# - 目录: 8条
# - 一级标题: 7条 × 6个 = 42条
# - 二级标题: 6条 × 21个 = 126条（如果全部检查）
# - ...
# 总计约200-300条规则
```
---

# PART B: 给 Agent AI 的提示词

> 此提示词将直接复制给另一个AI用于验证检查结果

````markdown
# 浙江师范大学本科毕业论文格式检测

## 角色
你是浙江师范大学本科毕业论文格式检测审核专家。

## 任务
验证脚本生成的格式检查结果，确认完成度计算是否准确，生成最终报告。

---

## 输入1：格式规范

/Users/CodeProjects/OfficeTest/浙江师范大学本科毕业论文（设计）格式规范_filtered.md

---

## 输入2：格式检查结果

```json
<此处粘贴脚本生成的 format_check_vXX.json>
```

**JSON说明**：
- `rule_checks`: 脚本已完成的所有规则检查结果（200-300条规则）
- `completion_summary`: 脚本预计算的完成度（按类别和总体）
- `failed_rules_detail`: 脚本提取的失败规则详情

**脚本已完成的工作**：
1. 单位换算和数据提取
2. 对照内置规则库，逐条检查格式（每条规则都有passed/failed状态）
3. 计算每个类别的完成度
4. 计算总体完成度
5. 提取失败规则的详细信息

---

## 你的工作

### 步骤1：验证规则判断

抽查部分规则检查结果，验证脚本的判断是否合理：

**抽查重点**：
- 失败的规则（`passed: false`）：确认actual和expected确实不匹配
- 关键格式项：页边距、一级标题、二级标题等核心格式
- 边界情况：如字号"三号(16pt)"和"小三(15pt)"的区分

**示例验证**：
```json
{
  "rule_id": "R050",
  "rule_name": "一级标题-字体",
  "expected": "宋体",
  "actual": "Times New Roman",
  "passed": false
}
```
验证：expected要求宋体，actual确实是Times New Roman，判断正确 ✓

### 步骤2：验证完成度计算

检查脚本的完成度计算：

**验证公式**：
```
类别完成度 = 通过规则数 / 总规则数 × 100%
```

**示例验证**：
```json
"page_setup": {
  "completion": {
    "total": 6,
    "passed": 5,
    "failed": 1,
    "completion": 83.3
  }
}
```
验证：5/6 = 83.33%，四舍五入为83.3% ✓

**总体完成度验证**：
```
总体 = Σ(类别完成度 × 权重)
```
对照权重分配，验证overall值是否正确。

### 步骤3：补充语义判断

脚本无法处理的主观要求（需要你判断）：
- 标题是否"简明扼要，无标点符号"
- 关键词数量是否3-5个
- 其他规范中的主观描述

### 步骤4：生成最终报告

如果脚本检查结果正确，直接使用其数据生成报告：

```json
{
  "overall_completion": 91.2,
  "completion_by_category": {
    "页面设置": 83.3,
    "中文摘要": 100.0,
    "英文摘要": 100.0,
    "目录": 90.0,
    "正文格式": 82.5,
    "图片": 100.0,
    "表格": 95.0,
    "参考文献": 100.0,
    "页眉页脚": 100.0,
    "致谢": 100.0,
    "附录": 100.0
  },
  "failed_rules": [
    {
      "category": "页面设置",
      "rule": "页边距-下",
      "expected": "2.0 cm",
      "actual": "1.8 cm",
      "severity": "medium",
      "suggestion": "调整下边距为2.0 cm"
    },
    {
      "category": "正文-一级标题",
      "rule": "一级标题-字体",
      "expected": "宋体",
      "actual": "Times New Roman",
      "locations": ["第1章 绪论", "第2章 深度学习与图像识别基础"],
      "severity": "critical",
      "suggestion": "将一级标题字体改为宋体"
    }
  ],
  "summary": {
    "total_issues": 22,
    "critical_issues": 6,
    "medium_issues": 1,
    "low_issues": 15
  },
  "recommendations": [
    "严重问题：6个一级标题字体错误，需改为宋体",
    "中等问题：下边距偏小0.2cm",
    "整体评价：文档格式完成度91.2%，修正22处问题即可达到95%以上"
  ]
}
```

**如果发现脚本错误**：
- 在报告中说明哪些规则的判断有误
- 提供修正后的完成度

---

## 注意事项

1. **信任脚本为主**：脚本已做了详细检查，除非发现明显错误，否则采纳其判断
2. **抽查验证**：不需要检查所有200+条规则，抽查关键项和失败项即可
3. **补充主观判断**：关注脚本无法处理的主观要求
4. **简洁报告**：利用脚本提供的数据，生成清晰的报告
5. **问题分级**：使用脚本已提供的severity分级

现在开始验证和生成报告。
````
