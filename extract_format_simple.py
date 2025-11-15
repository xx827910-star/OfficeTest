#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
粗粒度论文格式检测方案 - 格式数据提取脚本

功能：
1. 读取 batch_output/ 中的 JSON 文件
2. 提取格式信息（只提取原始数据，不做格式检查）
3. 输出简洁的 format_data_vXX.json

设计理念：
- 脚本只负责数据提取和单位转换
- AI 负责所有格式判断、规范对比和完成度计算
"""

import json
import math
import os
import re
from collections import Counter
from pathlib import Path
from typing import Dict, List, Any, Optional


# ==================== 单位换算工具 ====================

def twips_to_cm(twips: str) -> str:
    """Twips → cm (1 cm = 567 twips)"""
    if not twips or twips == "":
        return ""
    try:
        cm = int(twips) / 567
        return f"{cm:.1f} cm"
    except (ValueError, TypeError):
        return ""


def half_point_to_pt_and_chinese(half_point: str) -> str:
    """半磅 → pt → 中文字号"""
    if not half_point or half_point == "":
        return ""

    SIZE_MAP = {
        16: "三号",
        15: "小三",
        14: "四号",
        12: "小四",
        10.5: "五号",
        9: "小五"
    }

    try:
        pt = int(half_point) / 2
        chinese = SIZE_MAP.get(pt, "")
        if chinese:
            return f"{chinese}({pt}pt)"
        else:
            return f"{pt}pt"
    except (ValueError, TypeError):
        return ""


def twips_to_line_spacing(twips: str) -> str:
    """行距 twips → 倍数"""
    if not twips or twips == "":
        return ""
    try:
        multiplier = int(twips) / 240
        return f"{multiplier:.1f}倍"
    except (ValueError, TypeError):
        return ""


def twips_to_pt(twips: str) -> str:
    """段前/段后间距 twips → 磅"""
    if not twips or twips == "":
        return ""
    try:
        pt = int(twips) / 20
        return f"{pt:.0f}磅"
    except (ValueError, TypeError):
        return ""


def twips_to_pt_precise(twips: str, decimals: int = 1) -> str:
    """Twips → pt（可选精度）"""
    if not twips or twips == "":
        return ""
    try:
        pt = int(twips) / 20
        return f"{pt:.{decimals}f}pt"
    except (ValueError, TypeError):
        return ""


def twips_to_chars(twips: str) -> str:
    """首行缩进 twips → 字符（小四12pt，1字符=240 twips）"""
    if not twips or twips == "":
        return ""
    try:
        chars = int(twips) / 240
        return f"{chars:.0f}字符"
    except (ValueError, TypeError):
        return ""


def twips_to_chars_for_toc(twips: str, font_size_pt: float = 10.5) -> str:
    """TOC条目左缩进 twips → 字符（五号10.5pt，1字符=210 twips）"""
    if not twips or twips == "":
        return ""
    try:
        twips_int = int(twips)
        twips_per_char = font_size_pt * 20  # 1pt = 20 twips
        chars = twips_int / twips_per_char
        return f"{chars:.1f}字符"
    except (ValueError, TypeError):
        return ""


def border_size_to_pt(size_value: str) -> str:
    """表格边框尺寸（1/8pt单位）→ pt"""
    if not size_value:
        return ""
    try:
        pt = int(size_value) / 8
        return f"{pt:.2f}pt"
    except (ValueError, TypeError):
        return ""


def is_blank_paragraph(para: Optional[Dict]) -> bool:
    if not para:
        return False
    return para.get('Text', '').strip() == ''


def format_paragraph_defaults(info: Dict[str, Any]) -> Dict[str, Any]:
    if not info:
        return {}
    return {
        "alignment": get_alignment(info.get('Alignment', '')),
        "line_spacing": twips_to_line_spacing(info.get('LineSpacing', '')),
        "spacing_before": twips_to_pt(info.get('SpacingBefore', '')),
        "spacing_after": twips_to_pt(info.get('SpacingAfter', '')),
        "first_line_indent": twips_to_chars(info.get('FirstLineIndent', '')),
        "first_line_indent_pt": twips_to_pt_precise(info.get('FirstLineIndent', '')),
    }


def format_run_defaults(info: Dict[str, Any]) -> Dict[str, Any]:
    if not info:
        return {}
    size = info.get('FontSize', '')
    return {
        "font_chinese": info.get('FontNameEastAsia', ''),
        "font_english": info.get('FontNameAscii', ''),
        "size": half_point_to_pt_and_chinese(size),
        "bold": info.get('Bold', False),
        "italic": info.get('Italic', False),
        "color": info.get('Color', ''),
    }


# ==================== 对齐方式映射 ====================

ALIGNMENT_MAP = {
    'center': '居中',
    'justificationvalues { }': '未提供',
    'left': '左对齐',
    'leftvalues { }': '左对齐',
    'start': '左对齐',
    'right': '右对齐',
    'rightvalues { }': '右对齐',
    'end': '右对齐',
    'both': '两端对齐',
    'bothvalues { }': '两端对齐',
    'distribute': '分散对齐',
    'thaidistribute': '分散对齐',
    'highkashida': 'HighKashida',
    'mediumkashida': 'MediumKashida',
    'lowkashida': 'LowKashida',
    # Math equation alignments
    'centergroup': '居中(组)',
    'centergroupvalues { }': '居中(组)',
    '': '未提供'
}


def get_alignment(alignment: str) -> str:
    """获取对齐方式的中文描述"""
    if alignment is None:
        alignment = ''
    key = alignment.strip()
    lower_key = key.lower()
    return ALIGNMENT_MAP.get(lower_key, key if key else '未提供')


# ==================== 样式继承解析 ====================

def build_styles_dict(styles: List[Dict]) -> Dict[str, Dict]:
    """将样式列表转换为字典，方便查找"""
    styles_dict = {}
    for style in styles:
        style_id = style.get('StyleId', '')
        if style_id:
            styles_dict[style_id] = style
    return styles_dict


def get_effective_font(para: Dict, styles_dict: Dict[str, Dict]) -> Dict[str, Any]:
    """
    解析最终生效的字体（处理继承）

    优先级：Run → StyleId → Normal → 默认值
    """
    # 默认值
    result = {
        'chinese': '宋体',
        'english': 'Times New Roman',
        'size': '12pt',
        'size_half_point': '24',
        'bold': False,
        'italic': False
    }

    # 从 StyleId 获取
    style_id = para.get('StyleId', '') or 'Normal'
    if style_id in styles_dict:
        style = styles_dict[style_id]
        run_props = style.get('RunProperties')
        if run_props:
            if run_props.get('FontNameEastAsia'):
                result['chinese'] = run_props['FontNameEastAsia']
            if run_props.get('FontNameAscii'):
                result['english'] = run_props['FontNameAscii']
            if run_props.get('FontSize'):
                result['size_half_point'] = run_props['FontSize']
                result['size'] = half_point_to_pt_and_chinese(run_props['FontSize'])
            if run_props.get('Bold') is not None:
                result['bold'] = run_props['Bold']
            if run_props.get('Italic') is not None:
                result['italic'] = run_props['Italic']

    # 从 Run 获取（优先级最高）
    runs = para.get('Runs', [])
    if runs:
        # 目录段落中的第一个 Run 通常是字段标记，没有字号信息；
        # 这里找出第一个真正包含字体或字号设置的 Run。
        run_with_size = next((r for r in runs if r.get('FontSize')), None)
        run_with_font = next(
            (r for r in runs if r.get('FontNameEastAsia') or r.get('FontNameAscii')),
            None
        )
        run = run_with_size or run_with_font

        if run:
            if run.get('FontNameEastAsia'):
                result['chinese'] = run['FontNameEastAsia']
            if run.get('FontNameAscii'):
                result['english'] = run['FontNameAscii']
            if run.get('FontSize'):
                result['size_half_point'] = run['FontSize']
                result['size'] = half_point_to_pt_and_chinese(run['FontSize'])
            if run.get('Bold') is not None:
                result['bold'] = run['Bold']
            if run.get('Italic') is not None:
                result['italic'] = run['Italic']

    return result


def summarize_paragraph_format(para: Dict, styles_dict: Dict[str, Dict], include_spacing: bool = False, is_toc: bool = False) -> Dict[str, Any]:
    """提炼段落的核心格式属性

    Args:
        para: 段落数据
        styles_dict: 样式字典
        include_spacing: 是否包含段前段后间距
        is_toc: 是否为TOC条目（会提取LeftIndent和TabStops）
    """
    font = get_effective_font(para, styles_dict)
    summary = {
        "index": para.get('Index'),
        "text": para.get('Text', '').strip(),
        "font": font['chinese'],
        "font_english": font['english'],
        "size": font['size'],
        "bold": font['bold'],
        "alignment": get_alignment(para.get('Alignment', '')),
        "first_line_indent": twips_to_chars(para.get('FirstLineIndent', '')),
        "first_line_indent_pt": twips_to_pt_precise(para.get('FirstLineIndent', '')),
    }

    line_spacing = para.get('LineSpacing', '')
    if line_spacing:
        summary["line_spacing"] = twips_to_line_spacing(line_spacing)

    if include_spacing:
        summary["spacing_before"] = twips_to_pt(para.get('SpacingBefore', ''))
        summary["spacing_after"] = twips_to_pt(para.get('SpacingAfter', ''))

    # TOC特殊处理：提取左缩进和制表位
    if is_toc:
        left_indent = para.get('LeftIndent', '')
        if left_indent:
            summary["left_indent"] = twips_to_chars_for_toc(left_indent)
            summary["left_indent_pt"] = twips_to_pt_precise(left_indent)
            summary["left_indent_twips"] = left_indent

            # 判断TOC层级
            try:
                indent_value = int(left_indent)
                if indent_value == 0 or indent_value == '':
                    summary["toc_level"] = 1
                elif indent_value == 210:
                    summary["toc_level"] = 2
                elif indent_value == 420:
                    summary["toc_level"] = 3
                else:
                    summary["toc_level"] = None  # 未知层级
            except (ValueError, TypeError):
                summary["toc_level"] = None
        else:
            summary["left_indent"] = ""
            summary["toc_level"] = 1  # 默认为一级（无缩进）

        # 提取制表位信息
        tab_stops = para.get('TabStops', [])
        if tab_stops:
            summary["tab_stops"] = []
            for tab in tab_stops:
                tab_info = {
                    "position": twips_to_cm(tab.get('Position', '')),
                    "position_pt": twips_to_pt_precise(tab.get('Position', '')),
                    "position_twips": tab.get('Position', ''),
                    "alignment": tab.get('Alignment', ''),
                    "leader": tab.get('Leader', ''),
                }
                summary["tab_stops"].append(tab_info)

    return summary


def normalize_tab_stops(tab_stops: Optional[List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """提取对齐判断需要的制表位核心字段"""
    normalized: List[Dict[str, Any]] = []
    if not tab_stops:
        return normalized
    for tab in tab_stops:
        normalized.append({
            "position": tab.get("position", ""),
            "alignment": tab.get("alignment", ""),
            "leader": tab.get("leader", "")
        })
    return normalized


def tab_stops_signature(tab_stops: Optional[List[Dict[str, Any]]]) -> tuple:
    """将制表位转换为可哈希的签名"""
    signature: List[tuple] = []
    if tab_stops:
        for tab in tab_stops:
            signature.append((
                tab.get("position_twips") or tab.get("position_pt") or tab.get("position", ""),
                tab.get("alignment", ""),
                tab.get("leader", "")
            ))
    return tuple(signature)


def aggregate_toc_items(
    items: List[Dict[str, Any]],
    sample_config: Optional[Dict[str, Any]] = None
) -> Dict[str, List[Dict[str, Any]]]:
    """将目录条目按格式 profile 聚合"""
    profiles: List[Dict[str, Any]] = []
    anomalies: List[Dict[str, Any]] = []
    if not items:
        return {"profiles": profiles, "anomalies": anomalies}

    profile_map: Dict[tuple, Dict[str, Any]] = {}

    for item in items:
        signature = (
            item.get('toc_level'),
            item.get('font', ''),
            item.get('size', ''),
            item.get('bold', False),
            item.get('alignment', ''),
            item.get('left_indent', ''),
            tab_stops_signature(item.get('tab_stops'))
        )

        if signature not in profile_map:
            profile_id = f"toc_profile_{len(profile_map) + 1}"
            profile_map[signature] = {
                "profile_id": profile_id,
                "toc_level": item.get('toc_level'),
                "font": item.get('font', ''),
                "size": item.get('size', ''),
                "bold": item.get('bold', False),
                "alignment": item.get('alignment', ''),
                "left_indent": item.get('left_indent', ''),
                "tab_stops": normalize_tab_stops(item.get('tab_stops')),
                "count": 0,
                "indexes": [],
                "sample_text": item.get('text', '')[:100]
            }

        profile_entry = profile_map[signature]
        profile_entry["count"] += 1
        profile_entry["indexes"].append(item.get('index'))
        if not profile_entry["sample_text"]:
            profile_entry["sample_text"] = item.get('text', '')[:100]

    profiles = list(profile_map.values())
    profiles.sort(key=lambda p: p["indexes"][0] if p["indexes"] else float('inf'))

    primary_profile = max(profiles, key=lambda p: p["count"]) if profiles else None

    anomaly_records: List[Dict[str, Any]] = []
    if primary_profile:
        for profile in profiles:
            if profile is primary_profile:
                continue
            differences = {}
            for field in ["toc_level", "font", "size", "bold", "alignment", "left_indent"]:
                if profile.get(field) != primary_profile.get(field):
                    differences[field] = profile.get(field)
            if profile.get("tab_stops") != primary_profile.get("tab_stops"):
                differences["tab_stops"] = profile.get("tab_stops")
            anomaly_records.append({
                "profile": profile,
                "differences": differences
            })

    if sample_config:
        apply_sampling_to_profiles(profiles, sample_config)

    for record in anomaly_records:
        profile = record["profile"]
        anomaly_entry = {
            "profile_id": profile["profile_id"],
            "count": profile.get("count", 0),
            "differences": record["differences"],
            "indexes": profile.get("indexes", [])
        }
        anomalies.append(anomaly_entry)

    return {"profiles": profiles, "anomalies": anomalies}


def aggregate_format_profiles(
    items: List[Dict[str, Any]],
    key_fields: List[str],
    profile_prefix: str,
    sample_config: Optional[Dict[str, Any]] = None
) -> Dict[str, List[Dict[str, Any]]]:
    """通用格式 profile 聚合，用于正文标题/正文段落"""
    profiles: List[Dict[str, Any]] = []
    deviations: List[Dict[str, Any]] = []
    if not items:
        return {"profiles": profiles, "deviations": deviations}

    profile_map: Dict[tuple, Dict[str, Any]] = {}

    for item in items:
        signature = tuple(item.get(field, '') for field in key_fields)
        if signature not in profile_map:
            profile_id = f"{profile_prefix}_{len(profile_map) + 1}"
            profile_map[signature] = {
                "profile_id": profile_id,
                "count": 0,
                "indexes": [],
                "sample_text": item.get('text', '')[:100] if item.get('text') else ""
            }
            for field in key_fields:
                profile_map[signature][field] = item.get(field, '')

        profile_entry = profile_map[signature]
        profile_entry["count"] += 1
        profile_entry["indexes"].append(item.get('index'))
        if not profile_entry["sample_text"]:
            profile_entry["sample_text"] = item.get('text', '')[:100]

    profiles = list(profile_map.values())
    profiles.sort(key=lambda p: p["indexes"][0] if p["indexes"] else float('inf'))

    primary_profile = max(profiles, key=lambda p: p["count"]) if profiles else None

    if primary_profile:
        for profile in profiles:
            if profile is primary_profile:
                continue
            differences = {}
            for field in key_fields:
                if profile.get(field) != primary_profile.get(field):
                    differences[field] = profile.get(field)
            deviations.append({
                "profile_id": profile["profile_id"],
                "indexes": profile.get("indexes", []),
                "differences": differences
            })

    if sample_config:
        apply_sampling_to_profiles(profiles, sample_config)

    return {"profiles": profiles, "deviations": deviations}


BODY_SAMPLING_CONFIG = {"ratio": 0.15, "min": 6, "max": 12}
TOC_SAMPLING_CONFIG = {"ratio": 0.2, "min": 3, "max": 8}


def compute_sample_size(total: int, config: Dict[str, Any]) -> int:
    if total <= 0:
        return 0
    ratio = config.get("ratio", 0)
    minimum = config.get("min", 1)
    maximum = config.get("max", total)
    if total <= minimum:
        return total
    size = max(minimum, math.ceil(total * ratio)) if ratio > 0 else minimum
    size = min(size, maximum, total)
    return size


def evenly_spaced_sample(values: List[int], sample_size: int) -> List[int]:
    cleaned = sorted(v for v in values if isinstance(v, int))
    n = len(cleaned)
    if n == 0 or sample_size <= 0:
        return []
    if sample_size >= n:
        return cleaned
    if sample_size == 1:
        return [cleaned[n // 2]]

    positions = []
    for i in range(sample_size):
        pos = round(i * (n - 1) / (sample_size - 1))
        positions.append(pos)

    seen = set()
    sampled = []
    for pos in positions:
        if pos not in seen:
            seen.add(pos)
            sampled.append(cleaned[pos])

    if len(sampled) < sample_size:
        for idx in range(n):
            if idx not in seen:
                seen.add(idx)
                sampled.append(cleaned[idx])
                if len(sampled) == sample_size:
                    break

    return sorted(sampled)


def sample_indexes(indexes: List[Any], config: Dict[str, Any]) -> List[int]:
    cleaned = [idx for idx in indexes if isinstance(idx, int)]
    total = len(cleaned)
    sample_size = compute_sample_size(total, config)
    if sample_size <= 0:
        return []
    return evenly_spaced_sample(cleaned, sample_size)


def apply_sampling_to_profiles(profiles: List[Dict[str, Any]], config: Dict[str, Any]) -> None:
    for profile in profiles:
        indexes = profile.get("indexes", [])
        profile["indexes"] = sample_indexes(indexes, config)


TABLE_CAPTION_FIELDS = [
    "font",
    "font_english",
    "size",
    "bold",
    "alignment",
    "first_line_indent",
    "first_line_indent_pt",
    "line_spacing",
    "spacing_before",
    "spacing_after",
    "blank_before",
    "blank_after"
]

TABLE_SOURCE_FIELDS = [
    "font",
    "font_english",
    "size",
    "bold",
    "alignment",
    "first_line_indent",
    "first_line_indent_pt",
    "line_spacing",
    "spacing_before",
    "spacing_after"
]


def most_common_value(values: List[Any]) -> Any:
    filtered = [v for v in values if v not in (None, "")]
    if not filtered:
        return ""
    counter = Counter(filtered)
    max_count = max(counter.values())
    for val in filtered:
        if counter[val] == max_count:
            return val
    return filtered[0]


def summarize_table_entries(entries: List[Dict[str, Any]]) -> Dict[str, Any]:
    summary = {
        "count": len(entries),
        "defaults": {
            "caption": {},
            "source": {}
        },
        "entries": [],
        "stats": {
            "with_source": 0,
            "without_source": 0
        }
    }

    if not entries:
        return summary

    for field in TABLE_CAPTION_FIELDS:
        summary["defaults"]["caption"][field] = most_common_value([
            entry["caption"].get(field) for entry in entries
        ])

    source_entries = [entry for entry in entries if entry.get("source")]
    summary["stats"]["with_source"] = len(source_entries)
    summary["stats"]["without_source"] = len(entries) - len(source_entries)

    if source_entries:
        for field in TABLE_SOURCE_FIELDS:
            summary["defaults"]["source"][field] = most_common_value([
                entry["source"].get(field) for entry in source_entries
            ])

    for entry in entries:
        caption_diff = {}
        for field in TABLE_CAPTION_FIELDS:
            value = entry["caption"].get(field)
            if value != summary["defaults"]["caption"].get(field):
                caption_diff[field] = value

        entry_summary: Dict[str, Any] = {
            "index": entry["index"],
            "text": entry["text"],
            "caption_diff": caption_diff
        }

        source_data = entry.get("source")
        if source_data:
            source_diff = {}
            for field in TABLE_SOURCE_FIELDS:
                value = source_data.get(field)
                if value != summary["defaults"]["source"].get(field):
                    source_diff[field] = value
            entry_summary["source"] = {
                "index": source_data.get("index"),
                "text": source_data.get("text", ""),
                "diff": source_diff
            }
        else:
            entry_summary["source"] = {"missing": True}

        summary["entries"].append(entry_summary)

    return summary


# ==================== 段落分类 ====================

def classify_paragraph(para: Dict) -> str:
    """
    段落类型分类

    返回值：
    - ABSTRACT_CN_TITLE: 中文摘要标题
    - ABSTRACT_EN_TITLE: 英文摘要标题
    - TOC_TITLE: 目录标题
    - TOC_REF: 目录项（含 PAGEREF）
    - REFERENCE_TITLE: 参考文献标题
    - ACKNOWLEDGEMENT_TITLE: 致谢标题
    - APPENDIX_TITLE: 附录标题
    - HEADING_1: 一级标题（第X章）
    - HEADING_2: 二级标题（X.X）
    - HEADING_3: 三级标题（X.X.X）
    - KEYWORDS_CN: 中文关键词
    - KEYWORDS_EN: 英文关键词
    - FIGURE_CAPTION: 图标题
    - TABLE_CAPTION: 表标题
    - REFERENCE_ITEM: 参考文献条目
    - BODY: 正文
    """
    text = para.get('Text', '').strip()

    # 空段落
    if not text:
        return 'EMPTY'

    # 目录项（含 PAGEREF）
    if 'PAGEREF' in text:
        return 'TOC_REF'

    # 特殊标题
    if text in ['摘  要', '摘要']:
        return 'ABSTRACT_CN_TITLE'
    if text == 'Abstract':
        return 'ABSTRACT_EN_TITLE'
    if text in ['目  录', '目录']:
        return 'TOC_TITLE'
    if text == '参考文献':
        return 'REFERENCE_TITLE'
    if text in ['致  谢', '致谢']:
        return 'ACKNOWLEDGEMENT_TITLE'
    if text in ['附  录', '附录']:
        return 'APPENDIX_TITLE'

    # 标题
    if re.match(r'^第\d+章\s+', text):
        return 'HEADING_1'
    if re.match(r'^\d+\.\d+\s+\S', text):
        return 'HEADING_2'
    if re.match(r'^\d+\.\d+\.\d+\s+\S', text):
        return 'HEADING_3'

    # 关键词
    if text.startswith('关键词：') or text.startswith('关键词:'):
        return 'KEYWORDS_CN'
    if text.startswith('Key words') or text.startswith('Keywords'):
        return 'KEYWORDS_EN'

    # 图表标题 - 使用 C# 提供的 HasCaptionField 布尔值判断
    if para.get('HasCaptionField'):
        caption_type = para.get('CaptionFieldType', '')
        if caption_type == 'Table':
            return 'TABLE_CAPTION'
        elif caption_type == 'Figure':
            return 'FIGURE_CAPTION'

    # 参考文献条目
    if re.match(r'^\[\d+\]', text):
        return 'REFERENCE_ITEM'

    return 'BODY'


# ==================== Caption文本清理和编号推断 ====================

def clean_caption_text(text: str, chapter: int, seq_num: int, caption_type: str) -> str:
    """
    清理Caption文本：移除域代码，推断完整编号

    Args:
        text: 原始文本，如 "表1- SEQ Table_1 \* ARABIC  主要深度学习..."
        chapter: 章节号，如 1
        seq_num: 该章内的序号，如 1
        caption_type: "Table" 或 "Figure"

    Returns:
        清理后的文本，如 "表1-1 主要深度学习..."
    """
    # 移除域代码部分 (SEQ xxx \* ARABIC)
    text = re.sub(r'\s*SEQ\s+\w+_?\w*\s*\\?\*?\s*\w*\s*', '', text)

    # 推断完整编号
    prefix = "表" if caption_type == "Table" else "图"

    # 检查是否已有完整编号
    if re.match(rf'^{prefix}\d+-\d+\s', text):
        # 已有完整编号，如 "表1-1 xxx"
        return text

    # 检查是否只有章节号（匹配 "表1-" 或 "表1 "）
    match = re.match(rf'^{prefix}(\d+)-?\s*(.*)', text)
    if match:
        # 提取章节号和剩余文本
        text_chapter = int(match.group(1))
        remaining_text = match.group(2).strip()

        # 构造完整编号
        return f"{prefix}{text_chapter}-{seq_num} {remaining_text}"

    # 如果完全没有编号（不太可能），添加编号
    return f"{prefix}{chapter}-{seq_num} {text}"


def infer_chapter_from_text(text: str) -> Optional[int]:
    """从Caption文本中提取章节号"""
    # 匹配 "表1-" 或 "图3-"
    match = re.match(r'^[表图](\d+)-', text)
    if match:
        return int(match.group(1))
    return None


# ==================== 格式数据提取 ====================

def extract_format_data(input_json_path: str) -> Dict[str, Any]:
    """
    从 JSON 文件中提取格式数据

    只输出格式数据（value），不输出 expected 和 match
    """
    with open(input_json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 构建样式字典
    styles_dict = build_styles_dict(data.get('Styles', []))

    # 提取页面设置
    sections = data.get('Sections', [])
    section = sections[0] if sections else {}

    page_setup = {
        "margin_top": twips_to_cm(section.get('MarginTop', '')),
        "margin_bottom": twips_to_cm(section.get('MarginBottom', '')),
        "margin_left": twips_to_cm(section.get('MarginLeft', '')),
        "margin_right": twips_to_cm(section.get('MarginRight', '')),
        "page_width": twips_to_cm(section.get('PageWidth', '')),
        "page_height": twips_to_cm(section.get('PageHeight', '')),
    }

    # 提取段落
    paragraphs = data.get('Paragraphs', [])
    paragraph_lookup = {
        para.get('Index'): para
        for para in paragraphs
        if para.get('Index') is not None
    }

    # 分类段落
    classified = {
        'ABSTRACT_CN_TITLE': [],
        'ABSTRACT_CN_CONTENT': [],
        'KEYWORDS_CN': [],
        'ABSTRACT_EN_TITLE': [],
        'ABSTRACT_EN_CONTENT': [],
        'KEYWORDS_EN': [],
        'TOC_TITLE': [],
        'TOC_REF': [],
        'HEADING_1': [],
        'HEADING_2': [],
        'HEADING_3': [],
        'BODY': [],
        'FIGURE_CAPTION': [],
        'TABLE_CAPTION': [],
        'REFERENCE_TITLE': [],
        'REFERENCE_ITEM': [],
        'ACKNOWLEDGEMENT_TITLE': [],
        'ACKNOWLEDGEMENT_CONTENT': [],
        'APPENDIX_TITLE': [],
        'APPENDIX_CONTENT': [],
    }

    # 状态标记
    in_abstract_cn = False
    in_abstract_en = False
    in_acknowledgement = False
    in_appendix = False
    last_figure_index: Optional[int] = None
    last_table_index: Optional[int] = None
    figure_sources: Dict[int, Dict] = {}
    table_sources: Dict[int, Dict] = {}

    for para in paragraphs:
        para_type = classify_paragraph(para)
        text = para.get('Text', '').strip()

        # 图表标题需要单独处理状态
        if para_type == 'FIGURE_CAPTION':
            classified['FIGURE_CAPTION'].append(para)
            last_figure_index = para.get('Index')
            last_table_index = None
            continue
        if para_type == 'TABLE_CAPTION':
            classified['TABLE_CAPTION'].append(para)
            last_table_index = para.get('Index')
            last_figure_index = None
            continue

        is_source_line = text.startswith('来源：') or text.startswith('来源:') or text.lower().startswith('source:')
        if is_source_line:
            if last_figure_index is not None:
                figure_sources[last_figure_index] = para
            elif last_table_index is not None:
                table_sources[last_table_index] = para
            continue

        # 状态机：处理多段内容
        if para_type == 'ABSTRACT_CN_TITLE':
            in_abstract_cn = True
            in_abstract_en = False
            in_acknowledgement = False
            in_appendix = False
            classified['ABSTRACT_CN_TITLE'].append(para)
        elif para_type == 'ABSTRACT_EN_TITLE':
            in_abstract_cn = False
            in_abstract_en = True
            in_acknowledgement = False
            in_appendix = False
            classified['ABSTRACT_EN_TITLE'].append(para)
        elif para_type == 'ACKNOWLEDGEMENT_TITLE':
            in_abstract_cn = False
            in_abstract_en = False
            in_acknowledgement = True
            in_appendix = False
            classified['ACKNOWLEDGEMENT_TITLE'].append(para)
        elif para_type == 'APPENDIX_TITLE':
            in_abstract_cn = False
            in_abstract_en = False
            in_acknowledgement = False
            in_appendix = True
            classified['APPENDIX_TITLE'].append(para)
        elif para_type == 'KEYWORDS_CN':
            in_abstract_cn = False
            classified['KEYWORDS_CN'].append(para)
        elif para_type == 'KEYWORDS_EN':
            in_abstract_en = False
            classified['KEYWORDS_EN'].append(para)
        elif para_type == 'BODY':
            if in_abstract_cn:
                classified['ABSTRACT_CN_CONTENT'].append(para)
            elif in_abstract_en:
                classified['ABSTRACT_EN_CONTENT'].append(para)
            elif in_acknowledgement:
                classified['ACKNOWLEDGEMENT_CONTENT'].append(para)
            elif in_appendix:
                classified['APPENDIX_CONTENT'].append(para)
            else:
                classified['BODY'].append(para)
        elif para_type in classified:
            classified[para_type].append(para)

    # ========== 清理Caption文本并推断完整编号 ==========
    # 为每个章节的表格和图片维护计数器
    table_counters = {}  # {chapter: count}
    figure_counters = {}  # {chapter: count}

    # 处理表格标题
    for para in classified['TABLE_CAPTION']:
        text = para.get('Text', '').strip()
        chapter = infer_chapter_from_text(text)

        if chapter is not None:
            # 更新计数器
            table_counters[chapter] = table_counters.get(chapter, 0) + 1
            seq_num = table_counters[chapter]

            # 清理文本并推断编号
            cleaned_text = clean_caption_text(text, chapter, seq_num, 'Table')
            para['Text'] = cleaned_text
            para['OriginalText'] = text  # 保留原始文本供调试

    # 处理图片标题
    for para in classified['FIGURE_CAPTION']:
        text = para.get('Text', '').strip()
        chapter = infer_chapter_from_text(text)

        if chapter is not None:
            # 更新计数器
            figure_counters[chapter] = figure_counters.get(chapter, 0) + 1
            seq_num = figure_counters[chapter]

            # 清理文本并推断编号
            cleaned_text = clean_caption_text(text, chapter, seq_num, 'Figure')
            para['Text'] = cleaned_text
            para['OriginalText'] = text  # 保留原始文本供调试

    # 构建输出结构
    result = {
        "page_setup": page_setup,
        "defaults": {
            "paragraph": format_paragraph_defaults(data.get('DefaultParagraphFormat', {})),
            "run": format_run_defaults(data.get('DefaultRunFormat', {})),
        },
        "sections": {}
    }

    section_settings = []
    for section_info in sections:
        section_settings.append({
            "index": section_info.get('Index'),
            "title_page": section_info.get('TitlePage'),
            "page_number_format": section_info.get('PageNumberFormat', ''),
            "page_number_start": section_info.get('PageNumberStart', ''),
            "header_references": section_info.get('HeaderReferences', []),
            "footer_references": section_info.get('FooterReferences', []),
        })

    if section_settings:
        result["sections"]["section_settings"] = section_settings

    # 中文摘要
    if classified['ABSTRACT_CN_TITLE']:
        title_para = classified['ABSTRACT_CN_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["abstract_cn"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
                "spacing_before": twips_to_pt(title_para.get('SpacingBefore', '')),
            },
            "content": {
                "count": len(classified['ABSTRACT_CN_CONTENT']),
                "items": []
            }
        }

        for para in classified['ABSTRACT_CN_CONTENT']:
            font = get_effective_font(para, styles_dict)
            result["sections"]["abstract_cn"]["content"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip()[:100] + "..." if len(para['Text'].strip()) > 100 else para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "first_line_indent": twips_to_chars(para.get('FirstLineIndent', '')),
                "alignment": get_alignment(para.get('Alignment', '')),
            })

        # 关键词
        if classified['KEYWORDS_CN']:
            kw_para = classified['KEYWORDS_CN'][0]
            text = kw_para['Text'].strip()
            # 提取关键词
            if '：' in text:
                keywords_part = text.split('：', 1)[1]
            elif ':' in text:
                keywords_part = text.split(':', 1)[1]
            else:
                keywords_part = text

            # 分隔符检测
            if '；' in keywords_part:
                separator = '；'
                keywords = keywords_part.split('；')
            elif ';' in keywords_part:
                separator = ';'
                keywords = keywords_part.split(';')
            else:
                separator = ''
                keywords = [keywords_part]

            font = get_effective_font(kw_para, styles_dict)
            result["sections"]["abstract_cn"]["keywords"] = {
                "text": text,
                "separator": separator,
                "keyword_count": len([k for k in keywords if k.strip()]),
                "label_bold": font['bold'],
                "first_line_indent": twips_to_chars(kw_para.get('FirstLineIndent', '')),
                "first_line_indent_pt": twips_to_pt_precise(kw_para.get('FirstLineIndent', '')),
            }

    # 英文摘要
    if classified['ABSTRACT_EN_TITLE']:
        title_para = classified['ABSTRACT_EN_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["abstract_en"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['english'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
                "spacing_before": twips_to_pt(title_para.get('SpacingBefore', '')),
            },
            "content": {
                "count": len(classified['ABSTRACT_EN_CONTENT']),
                "items": []
            }
        }

        for para in classified['ABSTRACT_EN_CONTENT']:
            font = get_effective_font(para, styles_dict)
            result["sections"]["abstract_en"]["content"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip()[:100] + "..." if len(para['Text'].strip()) > 100 else para['Text'].strip(),
                "font": font['english'],
                "size": font['size'],
                "first_line_indent": twips_to_chars(para.get('FirstLineIndent', '')),
                "alignment": get_alignment(para.get('Alignment', '')),
            })

        # 关键词
        if classified['KEYWORDS_EN']:
            kw_para = classified['KEYWORDS_EN'][0]
            text = kw_para['Text'].strip()
            # 提取关键词
            if ':' in text:
                keywords_part = text.split(':', 1)[1]
            else:
                keywords_part = text

            # 分隔符检测
            if ';' in keywords_part:
                separator = ';'
                keywords = keywords_part.split(';')
            elif ',' in keywords_part:
                separator = ','
                keywords = keywords_part.split(',')
            else:
                separator = ''
                keywords = [keywords_part]

            font = get_effective_font(kw_para, styles_dict)
            result["sections"]["abstract_en"]["keywords"] = {
                "text": text,
                "separator": separator,
                "keyword_count": len([k for k in keywords if k.strip()]),
                "label_bold": font['bold'],
                "first_line_indent": twips_to_chars(kw_para.get('FirstLineIndent', '')),
                "first_line_indent_pt": twips_to_pt_precise(kw_para.get('FirstLineIndent', '')),
            }

    # 目录
    if classified['TOC_TITLE']:
        title_para = classified['TOC_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["toc"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
                "spacing_before": twips_to_pt(title_para.get('SpacingBefore', '')),
            },
            "items_count": len(classified['TOC_REF']),
            "profiles": [],
            "anomalies": []
        }

        toc_items = []
        for para in classified['TOC_REF']:
            item = summarize_paragraph_format(para, styles_dict, include_spacing=True, is_toc=True)
            item["numbering_level"] = para.get('NumberingLevel', '')
            toc_items.append(item)

        toc_summary = aggregate_toc_items(toc_items, TOC_SAMPLING_CONFIG)
        result["sections"]["toc"]["profiles"] = toc_summary["profiles"]
        result["sections"]["toc"]["anomalies"] = toc_summary["anomalies"]

    # 正文 - 一级标题
    if classified['HEADING_1']:
        result["sections"]["main"] = {}
        h1_items = []
        for para in classified['HEADING_1']:
            font = get_effective_font(para, styles_dict)
            h1_items.append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
                "spacing_before": twips_to_pt(para.get('SpacingBefore', '')),
            })

        h1_summary = aggregate_format_profiles(
            h1_items,
            ["font", "size", "bold", "alignment", "spacing_before"],
            "h1_profile"
        )
        result["sections"]["main"]["h1"] = {
            "count": len(classified['HEADING_1']),
            "profiles": h1_summary["profiles"],
            "deviations": h1_summary["deviations"]
        }

    # 正文 - 二级标题
    if classified['HEADING_2']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        h2_items = []
        for para in classified['HEADING_2']:
            font = get_effective_font(para, styles_dict)
            h2_items.append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })
        h2_summary = aggregate_format_profiles(
            h2_items,
            ["font", "size", "bold", "alignment"],
            "h2_profile"
        )
        result["sections"]["main"]["h2"] = {
            "count": len(classified['HEADING_2']),
            "profiles": h2_summary["profiles"],
            "deviations": h2_summary["deviations"]
        }

    # 正文 - 三级标题
    if classified['HEADING_3']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        h3_items = []
        for para in classified['HEADING_3']:
            font = get_effective_font(para, styles_dict)
            h3_items.append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })
        h3_summary = aggregate_format_profiles(
            h3_items,
            ["font", "size", "bold", "alignment"],
            "h3_profile"
        )
        result["sections"]["main"]["h3"] = {
            "count": len(classified['HEADING_3']),
            "profiles": h3_summary["profiles"],
            "deviations": h3_summary["deviations"]
        }

    # 正文段落
    if classified['BODY']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        body_items = []
        for para in classified['BODY']:
            font = get_effective_font(para, styles_dict)
            line_spacing = twips_to_line_spacing(para.get('LineSpacing', ''))
            body_items.append({
                "index": para['Index'],
                "text": para['Text'].strip()[:100] + "..." if len(para['Text'].strip()) > 100 else para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "first_line_indent": twips_to_chars(para.get('FirstLineIndent', '')),
                "line_spacing": line_spacing,
                "alignment": get_alignment(para.get('Alignment', '')),
            })
        body_summary = aggregate_format_profiles(
            body_items,
            ["font", "size", "first_line_indent", "line_spacing", "alignment"],
            "body_profile",
            BODY_SAMPLING_CONFIG
        )
        result["sections"]["main"]["body"] = {
            "count": len(classified['BODY']),
            "profiles": body_summary["profiles"],
            "deviations": body_summary["deviations"]
        }

    # 图标题
    if classified['FIGURE_CAPTION']:
        result["sections"]["figures"] = {
            "count": len(classified['FIGURE_CAPTION']),
            "items": []
        }

        for para in classified['FIGURE_CAPTION']:
            figure_summary = summarize_paragraph_format(para, styles_dict, include_spacing=True)
            caption_index = para.get('Index')
            if isinstance(caption_index, int):
                prev_para = paragraph_lookup.get(caption_index - 1)
                next_para = paragraph_lookup.get(caption_index + 1)
                figure_summary["blank_before"] = is_blank_paragraph(prev_para)
                figure_summary["blank_after"] = is_blank_paragraph(next_para)
            else:
                figure_summary["blank_before"] = False
                figure_summary["blank_after"] = False

            source_para = figure_sources.get(caption_index) if isinstance(caption_index, int) else None
            if source_para:
                figure_summary["source"] = summarize_paragraph_format(source_para, styles_dict, include_spacing=True)

            result["sections"]["figures"]["items"].append(figure_summary)

    # 表标题
    if classified['TABLE_CAPTION']:
        table_entries: List[Dict[str, Any]] = []

        for para in classified['TABLE_CAPTION']:
            table_summary = summarize_paragraph_format(para, styles_dict, include_spacing=True)
            caption_index = para.get('Index')
            blank_before = False
            blank_after = False
            if isinstance(caption_index, int):
                prev_para = paragraph_lookup.get(caption_index - 1)
                next_para = paragraph_lookup.get(caption_index + 1)
                blank_before = is_blank_paragraph(prev_para)
                blank_after = is_blank_paragraph(next_para)

            caption_data = {
                field: table_summary.get(field, '')
                for field in TABLE_CAPTION_FIELDS
                if field not in {"blank_before", "blank_after"}
            }
            caption_data["blank_before"] = blank_before
            caption_data["blank_after"] = blank_after

            source_entry = None
            source_para = table_sources.get(caption_index) if isinstance(caption_index, int) else None
            if source_para:
                source_summary = summarize_paragraph_format(source_para, styles_dict, include_spacing=True)
                source_entry = {field: source_summary.get(field, '') for field in TABLE_SOURCE_FIELDS}
                source_entry["index"] = source_para.get('Index')
                source_entry["text"] = source_summary.get('text', '')

            table_entries.append({
                "index": caption_index,
                "text": table_summary.get('text', ''),
                "caption": caption_data,
                "source": source_entry
            })

        result["sections"]["tables"] = summarize_table_entries(table_entries)

        # 表格结构信息
        table_structures = []
        for table in data.get('Tables', []):
            top_border = table.get('TopBorder') or {}
            bottom_border = table.get('BottomBorder') or {}
            inside_h = table.get('InsideHorizontalBorder') or {}
            inside_v = table.get('InsideVerticalBorder') or {}
            table_structures.append({
                "index": table.get('Index'),
                "style_id": table.get('StyleId', ''),
                "alignment": get_alignment(table.get('Alignment', '')),
                "top_border": {
                    "style": top_border.get('Style', ''),
                    "size": border_size_to_pt(top_border.get('Size', '')),
                },
                "bottom_border": {
                    "style": bottom_border.get('Style', ''),
                    "size": border_size_to_pt(bottom_border.get('Size', '')),
                },
                "inside_horizontal": {
                    "style": inside_h.get('Style', ''),
                    "size": border_size_to_pt(inside_h.get('Size', '')),
                },
                "inside_vertical": {
                    "style": inside_v.get('Style', ''),
                    "size": border_size_to_pt(inside_v.get('Size', '')),
                },
                "has_inside_vertical": table.get('HasInsideVerticalBorders'),
                "has_vertical_outer": table.get('HasVerticalOuterBorders'),
                "has_inside_horizontal": table.get('HasInsideHorizontalBorders'),
            })

        if table_structures:
            result["sections"]["tables"]["structure"] = table_structures

    # 公式信息
    formulas = data.get('Formulas', [])
    if formulas:
        items = []
        for formula in formulas:
            para_index = formula.get('ParagraphIndex')
            para = paragraph_lookup.get(para_index)
            text_preview = ''

            # 提取公式字体和对齐方式
            equation_font = formula.get('EquationFont', '')
            equation_font_size = formula.get('EquationFontSize', '')
            alignment = get_alignment(formula.get('Alignment', ''))

            # 如果公式字体为空，从段落的Run中获取
            if para and (not equation_font or not equation_font_size):
                # 获取段落文本
                text_preview = para.get('Text', '').strip()

                # 从段落的Runs中获取字体信息
                runs = para.get('Runs', [])
                if runs:
                    # 找到第一个包含字体信息的Run
                    for run in runs:
                        if not equation_font:
                            font_ascii = run.get('FontNameAscii', '')
                            font_east_asia = run.get('FontNameEastAsia', '')
                            equation_font = font_east_asia or font_ascii
                        if not equation_font_size:
                            equation_font_size = run.get('FontSize', '')
                        if equation_font and equation_font_size:
                            break

                # 如果对齐方式为空，从段落属性中获取
                if not alignment or alignment == '未提供':
                    alignment = get_alignment(para.get('Alignment', ''))
            elif para:
                text_preview = para.get('Text', '').strip()

            items.append({
                "paragraph_index": para_index,
                "alignment": alignment,
                "numbering_text": formula.get('NumberingText', ''),
                "numbering_font": formula.get('NumberingFont', ''),
                "numbering_font_size": half_point_to_pt_and_chinese(formula.get('NumberingFontSize', '')),
                "equation_font": equation_font,
                "equation_font_size": half_point_to_pt_and_chinese(equation_font_size),
                "paragraph_preview": text_preview[:80] + '...' if len(text_preview) > 80 else text_preview,
            })

        result["sections"]["formulas"] = {
            "count": len(formulas),
            "items": items
        }

    # 参考文献
    if classified['REFERENCE_TITLE']:
        title_para = classified['REFERENCE_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["references"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
            },
            "items_count": len(classified['REFERENCE_ITEM']),
            "items": []
        }

        for para in classified['REFERENCE_ITEM'][:5]:  # 只取前5个
            font = get_effective_font(para, styles_dict)
            text = para['Text'].strip()
            bracket_match = re.match(r'^\[(\d+)\]', text)
            result["sections"]["references"]["items"].append({
                "index": para['Index'],
                "text": text[:150] + "..." if len(text) > 150 else text,
                "font": font['chinese'],
                "size": font['size'],
                "hanging_indent": twips_to_chars(para.get('HangingIndent', '')),
                "hanging_indent_pt": twips_to_pt_precise(para.get('HangingIndent', '')),
                "starts_with_bracket": bool(bracket_match),
                "sequence_number": int(bracket_match.group(1)) if bracket_match else None,
                "ends_with_period": text.endswith('。') or text.endswith('．') or text.endswith('.'),
            })

    # 页眉页脚
    headers = data.get('Headers', [])
    footers = data.get('Footers', [])

    def summarize_header_footer(items: List[Dict]) -> List[Dict]:
        summarized = []
        for item in items:
            paragraphs = item.get('Paragraphs', [])
            paragraph_formats = [
                summarize_paragraph_format(para, styles_dict, include_spacing=True)
                for para in paragraphs
            ]
            summarized.append({
                "index": item.get('Index'),
                "text": item.get('Text', '').strip(),
                "paragraphs": paragraph_formats
            })
        return summarized

    result["sections"]["headers_footers"] = {
        "header_count": len(headers),
        "footer_count": len(footers),
        "headers": summarize_header_footer(headers),
        "footers": summarize_header_footer(footers),
    }

    # 致谢
    if classified['ACKNOWLEDGEMENT_TITLE']:
        title_para = classified['ACKNOWLEDGEMENT_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["acknowledgement"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
                "spacing_before": twips_to_pt(title_para.get('SpacingBefore', '')),
            },
            "content_count": len(classified['ACKNOWLEDGEMENT_CONTENT']),
            "content_samples": [
                summarize_paragraph_format(para, styles_dict, include_spacing=True)
                for para in classified['ACKNOWLEDGEMENT_CONTENT'][:3]
            ]
        }

    # 附录
    if classified['APPENDIX_TITLE']:
        title_para = classified['APPENDIX_TITLE'][0]
        font = get_effective_font(title_para, styles_dict)
        result["sections"]["appendix"] = {
            "title": {
                "text": title_para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(title_para.get('Alignment', '')),
                "spacing_before": twips_to_pt(title_para.get('SpacingBefore', '')),
            },
            "content_count": len(classified['APPENDIX_CONTENT']),
            "content_samples": [
                summarize_paragraph_format(para, styles_dict, include_spacing=True)
                for para in classified['APPENDIX_CONTENT'][:3]
            ]
        }

    return result


# ==================== 主函数 ====================

def main():
    """主函数：批量处理所有版本的格式数据"""

    # 输入输出目录
    input_dir = Path('batch_output')
    output_dir = Path('json_output')  # 输出到 json_output 目录

    if not input_dir.exists():
        print(f"错误：输入目录 {input_dir} 不存在")
        return

    output_dir.mkdir(exist_ok=True)

    # 查找所有 JSON 文件（支持 v14_format_output.json 和 v01_xxx_format_output.json 两种格式）
    json_files = sorted(input_dir.glob('v*_format_output.json'))

    if not json_files:
        print(f"错误：在 {input_dir} 中没有找到任何 JSON 文件")
        return

    print(f"找到 {len(json_files)} 个 JSON 文件")
    print()

    # 处理每个文件
    for json_file in json_files:
        # 提取版本号（例如 v01）
        match = re.match(r'(v\d+)_', json_file.name)
        if not match:
            print(f"警告：无法从文件名 {json_file.name} 中提取版本号，跳过")
            continue

        version = match.group(1)
        output_file = output_dir / f'format_data_{version}.json'

        print(f"处理 {json_file.name} -> {output_file}")

        try:
            # 提取格式数据
            format_data = extract_format_data(str(json_file))

            # 写入输出文件
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(format_data, f, ensure_ascii=False, indent=2)

            print(f"  ✓ 成功生成 {output_file.name}")

        except Exception as e:
            print(f"  ✗ 错误：{e}")
            import traceback
            traceback.print_exc()

        print()

    print("批量处理完成！")


if __name__ == '__main__':
    main()
