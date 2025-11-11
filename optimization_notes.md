# JSON Slimming Discussions

## 1. TOC Profile Aggregation

- **现状问题**：`extract_format_simple.py:700-719` 会将目录里每个 `TOC_REF` 段落逐条写入 `sections.toc.items`，每条都携带 `font/size/bold/left_indent/tab_stops/...` 等重复字段。在 `format_data_v13.json` 中该块占 ~29 KB、约 2,000 行，信息高度冗余。
- **修改要点**：在构建 TOC 时改为“按格式 profile 聚合”。对每条 `TOC_REF` 计算键（如 `toc_level/font/size/left_indent/tab_positions/bold`），将相同 profile 的索引和示例文本收集到同一对象，并单独记 `anomalies` 列表存放与主 profile 不一致的条目。代码位置仍在 `extract_format_simple.py` 的 TOC 分支，只需新增 profile map 并替换原来的 `items` 列表输出。
- **修改后数据形态**：
  ```json
  "toc": {
    "title": { ... },
    "profiles": [
      {
        "profile_id": "L1_SongTi_5",
        "toc_level": 1,
        "font": "宋体",
        "size": "五号(10.5pt)",
        "left_indent": "0字符",
        "tab_stop_cm": ["15.5"],
        "count": 38,
        "indexes": [34, 45, 46, "..."],
        "sample_text": "第1章 绪论"
      }
    ],
    "anomalies": [
      {
        "index": 82,
        "profile_id": "L2_SongTi_5",
        "differences": { "tab_stop_cm": ["14.0"] }
      }
    ]
  }
  ```
  这样仍能指明每个条目的格式是否合规、对应段落索引在哪，只是把完全相同的字段折叠到 profile 上。AI 仍可据 `profiles` 判断主流格式、据 `anomalies` 精确定位不合规条目，因此不影响判断准确性；同时数据量可从 ~29 KB 直接降到 4–6 KB（profile 数通常 ≤4 个）。

## 2. 正文与标题格式去重

- **现状问题**：`extract_format_simple.py:723-803` 针对 `h1/h2/h3/body` 分别把每个段落的 `font/size/bold/alignment/...` 全量写入 `items`，即使一百个正文段落完全相同也要重复一百遍。在 `format_data_v13.json` 中，`sections.main` 占 ~4 KB，但随着正文段落数增长会线性膨胀，且大量重复信息让 AI 更难聚焦真正的异常。
- **修改要点**：与 TOC 类似，为正文/标题建立“格式 profile”。遍历段落时根据 `font/size/bold/first_line_indent/line_spacing/alignment` 组装键，把共享格式的段落索引聚合在 `profiles` 态中；另外记录 `deviations` 数组，仅当某段落字段与默认 profile 不同才写出差异。实现也在 `extract_format_simple.py` 的正文部分，通过一个 `Dictionary<string, Profile>` 保存聚合结果并在 JSON 输出时替换 `items`。
- **修改后数据形态**：
  ```json
  "main": {
    "h1": {
      "profiles": [
        {
          "profile_id": "H1_std",
          "font": "宋体",
          "size": "三号(16pt)",
          "bold": true,
          "alignment": "居中",
          "spacing_before": "12磅",
          "count": 6,
          "indexes": [22, 45, 78, 101, 134, 167],
          "sample_text": "第1章 绪论"
        }
      ],
      "deviations": [
        {
          "index": 190,
          "differences": { "alignment": "左对齐" }
        }
      ]
    },
    "body": {
      "profiles": [
        {
          "profile_id": "Body_std",
          "font": "宋体",
          "size": "小四(12pt)",
          "first_line_indent": "2字符",
          "line_spacing": "1.5倍",
          "alignment": "两端对齐",
          "count": 138,
          "indexes": [5, 6, 7, "..."],
          "sample_text": "随着深度学习技术的快速发展..."
        }
      ],
      "deviations": []
    }
  }
  ```
  这种结构仍让 AI 明确：标准 profile 是否符合规范？哪几个段落不符合？因为每个 profile 清晰列出属性，`deviations` 列出具体差异，并提供段落索引/示例文本，决策所需信息完全保留。相比原始逐条列举，数据量与 profile 数相关，而不是与段落总数相关，可在长篇正文里节约数十倍行数。

## 3. 按类别裁剪字段 *(暂不采纳)*

- **状态**：考虑到本阶段更关注结构性压缩，而非逐字段裁剪，此方案暂不采纳，留作备选。若后续仍需进一步瘦身，可回头讨论精简字段的安全范围及实现方案。

## 4. 共享字典 + 短 key

- **现状问题**：在 `format_data_v13.json` 中，大部分字符串（如 `"宋体"`, `"小四(12.0pt)"`, `"两端对齐"`) 在多个板块中重复数十次，单个字段名称也较长（`"first_line_indent"` 21 个字符）。这些重复纯属“描述同一枚值”，并非逻辑必需。
- **修改要点**：在 JSON 顶层新增 `lookups` 节点，集中列出常见枚举值与字段别名。例如：
  ```json
  "lookups": {
    "fonts": ["宋体", "Times New Roman"],
    "sizes": ["三号(16pt)", "小四(12pt)", "五号(10.5pt)"],
    "alignments": ["居中", "两端对齐"],
    "keys": { "font": "f", "size": "sz", "bold": "b", "alignment": "al", "indexes": "idx" }
  }
  ```
  数据体里改用索引与短 key：`{"f":0,"sz":2,"b":true,"al":1,"idx":[12,34,58]}`。实现层面只需在 `extract_format_simple.py` 尾部扫描收集去重值，并在序列化前把对象字段名映射到短 key（可用简单的 `dict` 转换）。
- **修改后数据形态**：
  ```json
  {
    "lookups": {
      "fonts": ["宋体", "Times New Roman"],
      "sizes": ["三号(16pt)", "五号(10.5pt)"],
      "keys": { "font": "f", "size": "sz", "alignment": "al", "indexes": "idx" }
    },
    "sections": {
      "toc": {
        "profiles": [
          { "al": 0, "f": 0, "sz": 1, "idx": [34,45,46] }
        ]
      }
    }
  }
  ```
  AI 在 prompt 中收到 `lookups.keys` 之后即可还原含义（例如“字段 `f` 指代字体，其值是 `lookups.fonts[f]`”）。信息准确度不受影响，因为所有原始值仍存在，只是去掉了重复写法。根据 `format_data_v13.json` 的重复统计，目录和正文的字符串占比约 60%，使用共享字典与短 key 可额外节省 ~10–15 KB。

## 5. 图/表差异化结构

- **现状问题**：`sections.figures.items` 与 `sections.tables.items` 中每条记录都包含 `font/size/alignment/blank_before/blank_after/source/...`，即使完全合规也要写全套字段；这让“没有问题的条目”占据绝大多数行数。例如 v13 文档 8 张图只有 1 张不合规，但 JSON 依旧为全部 8 张输出 ~150 行。
- **修改要点**：改成“默认值 + 仅记录偏差”。可以在相应分支先统计最常见的格式作为 `defaults`，输出一次；随后遍历每条图/表时，只在属性不同于默认值时写入 `differences`，否则只追加索引到 `conform_indexes`。实现上可让结构类似：
  ```json
  "figures": {
    "defaults": { "font": "宋体", "size": "五号(10.5pt)", "alignment": "居中", "blank_before": true, "blank_after": true },
    "conform_indexes": [120, 145, 167, 188, 210, 233, 255],
    "violations": [
      { "index": 178, "differences": { "blank_after": false } }
    ]
  }
  ```
  表格同理，还可以在 `violations` 内嵌入来源字号/三线表规格等信息。这样“没有问题”的条目只占一行（索引），真正的异常才写详情，极大减小体积。
- **准确性说明**：AI 判断是否符合规范所需的信息依然齐全——它可以从 `defaults` 知道应当字体=宋体五号、需要空行等；从 `conform_indexes` 得知这些段落都满足默认格式，不需逐条检查；从 `violations` 精确获取异常字段并生成 `failed_rules`。因此不会丢失任何判定所需的细节。

## 6. Prompt 分块投喂（流程优化）

- **现状问题**：即便 JSON 做了多轮瘦身，`prompt.md` 仍要求把整份规范 (~20 KB) + 完整 JSON (~几十 KB) 一次性交给 AI，大模型在输入超过 60 KB 时容易响应不稳定或中途截断。
- **修改要点**：在调用阶段把输入拆成多条消息：  
  1. 第 1 条发送角色设定与规范全文，请 AI 只阅读规范并回复“已了解”。  
  2. 后续依类别发送 JSON（如 `页面设置+摘要`, `目录`, `正文`, `图表+附录`），每块 <10 KB。  
  3. 最后一条以系统或用户消息要求：“你已经收到 4/4 块数据，请根据全部信息输出结果 JSON。”  
- **示例**：
  ```
  消息1：规范全文
  消息2：`{"part":1,"category":"页面设置+摘要","data":{...}}`
  消息3：`{"part":2,"category":"目录","data":{...}}`
  消息4：`{"part":3,"category":"正文","data":{...}}`
  消息5：`{"part":4,"category":"图表/附录","data":{...}}`
  消息6：最终任务提示
  ```
- **准确性说明**：分块只改变交互方式，不改变 JSON 内容；AI 仍能在最后一步同时参考所有信息（它的对话记忆里已经包含之前各块），不会影响判断。但由于每次输入更短，模型更稳定，且易于在出错时重新发送某一块。
