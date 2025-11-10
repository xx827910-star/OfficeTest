#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
方案A：粗粒度论文格式检测方案 - 格式数据提取脚本

功能：
1. 读取 batch_output/ 中的 JSON 文件
2. 提取格式信息（只提取原始数据，不做格式检查）
3. 输出简洁的 format_data_vXX.json

设计理念：
- 脚本只负责数据提取和单位转换
- AI 负责所有格式判断、规范对比和完成度计算
"""

import json
import os
import re
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


def twips_to_chars(twips: str) -> str:
    """首行缩进 twips → 字符"""
    if not twips or twips == "":
        return ""
    try:
        chars = int(twips) / 240
        return f"{chars:.0f}字符"
    except (ValueError, TypeError):
        return ""


# ==================== 对齐方式映射 ====================

ALIGNMENT_MAP = {
    'JustificationValues { }': '居中',
    'LeftValues { }': '左对齐',
    'RightValues { }': '右对齐',
    'BothValues { }': '两端对齐',
    '': '顶格/默认'
}


def get_alignment(alignment: str) -> str:
    """获取对齐方式的中文描述"""
    return ALIGNMENT_MAP.get(alignment, alignment)


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
        run = runs[0]  # 取第一个 Run
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

    # 图表标题
    if re.match(r'^图\d+-\d+', text):
        return 'FIGURE_CAPTION'
    if re.match(r'^表\d+-\d+', text):
        return 'TABLE_CAPTION'

    # 参考文献条目
    if re.match(r'^\[\d+\]', text):
        return 'REFERENCE_ITEM'

    return 'BODY'


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

    for para in paragraphs:
        para_type = classify_paragraph(para)

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

    # 构建输出结构
    result = {
        "page_setup": page_setup,
        "sections": {}
    }

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
            },
            "items_count": len(classified['TOC_REF']),
        }

    # 正文 - 一级标题
    if classified['HEADING_1']:
        result["sections"]["main"] = {}
        result["sections"]["main"]["h1"] = {
            "count": len(classified['HEADING_1']),
            "items": []
        }

        for para in classified['HEADING_1']:
            font = get_effective_font(para, styles_dict)
            result["sections"]["main"]["h1"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
                "spacing_before": twips_to_pt(para.get('SpacingBefore', '')),
            })

    # 正文 - 二级标题
    if classified['HEADING_2']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        result["sections"]["main"]["h2"] = {
            "count": len(classified['HEADING_2']),
            "items": []
        }

        for para in classified['HEADING_2'][:5]:  # 只取前5个
            font = get_effective_font(para, styles_dict)
            result["sections"]["main"]["h2"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })

    # 正文 - 三级标题
    if classified['HEADING_3']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        result["sections"]["main"]["h3"] = {
            "count": len(classified['HEADING_3']),
            "items": []
        }

        for para in classified['HEADING_3'][:5]:  # 只取前5个
            font = get_effective_font(para, styles_dict)
            result["sections"]["main"]["h3"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "bold": font['bold'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })

    # 正文段落
    if classified['BODY']:
        if "main" not in result["sections"]:
            result["sections"]["main"] = {}
        result["sections"]["main"]["body"] = {
            "count": len(classified['BODY']),
            "items": []
        }

        for para in classified['BODY'][:10]:  # 只取前10个
            font = get_effective_font(para, styles_dict)
            line_spacing = twips_to_line_spacing(para.get('LineSpacing', ''))
            result["sections"]["main"]["body"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip()[:100] + "..." if len(para['Text'].strip()) > 100 else para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "first_line_indent": twips_to_chars(para.get('FirstLineIndent', '')),
                "line_spacing": line_spacing,
                "alignment": get_alignment(para.get('Alignment', '')),
            })

    # 图标题
    if classified['FIGURE_CAPTION']:
        result["sections"]["figures"] = {
            "count": len(classified['FIGURE_CAPTION']),
            "items": []
        }

        for para in classified['FIGURE_CAPTION']:
            font = get_effective_font(para, styles_dict)
            result["sections"]["figures"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })

    # 表标题
    if classified['TABLE_CAPTION']:
        result["sections"]["tables"] = {
            "count": len(classified['TABLE_CAPTION']),
            "items": []
        }

        for para in classified['TABLE_CAPTION']:
            font = get_effective_font(para, styles_dict)
            result["sections"]["tables"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "alignment": get_alignment(para.get('Alignment', '')),
            })

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
            result["sections"]["references"]["items"].append({
                "index": para['Index'],
                "text": para['Text'].strip()[:150] + "..." if len(para['Text'].strip()) > 150 else para['Text'].strip(),
                "font": font['chinese'],
                "size": font['size'],
                "hanging_indent": twips_to_chars(para.get('HangingIndent', '')),
            })

    # 页眉页脚
    headers = data.get('Headers', [])
    footers = data.get('Footers', [])

    result["sections"]["headers_footers"] = {
        "header_count": len(headers),
        "footer_count": len(footers),
        "headers": [{"index": h.get('Index'), "text": h.get('Text', '').strip()} for h in headers],
        "footers": [{"index": f.get('Index'), "text": f.get('Text', '').strip()} for f in footers],
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
            },
            "content_count": len(classified['ACKNOWLEDGEMENT_CONTENT']),
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
            },
            "content_count": len(classified['APPENDIX_CONTENT']),
        }

    return result


# ==================== 主函数 ====================

def main():
    """主函数：批量处理所有版本的格式数据"""

    # 输入输出目录
    input_dir = Path('batch_output')
    output_dir = Path('.')  # 输出到当前目录

    if not input_dir.exists():
        print(f"错误：输入目录 {input_dir} 不存在")
        return

    # 查找所有 JSON 文件
    json_files = sorted(input_dir.glob('v*_*_format_output.json'))

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

        print(f"处理 {json_file.name} -> {output_file.name}")

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
