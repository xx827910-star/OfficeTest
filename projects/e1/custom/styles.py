"""E1 专用样式管理器。"""
import json
from typing import Any, Dict

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class E1StyleManager:
    """管理 E1 学校格式所需的样式参数。"""

    def __init__(self, config_source: Any):
        if isinstance(config_source, str):
            with open(config_source, 'r', encoding='utf-8') as f:
                self.config: Dict[str, Any] = json.load(f)
        else:
            self.config = dict(config_source or {})
        self.size_map = self.config.get('sizes', {})

    def get_font_size(self, size_name):
        """
        获取字号对应的磅值
        :param size_name: 字号名称（如"小四"）
        :return: Pt对象
        """
        size = self.size_map.get(size_name, 12)
        return Pt(size)

    def _resolve_size_value(self, style_config, default=12):
        if not style_config:
            return default
        if isinstance(style_config, (int, float)):
            return style_config
        if 'size' in style_config:
            return style_config['size']
        if 'size_name' in style_config:
            return self.size_map.get(style_config['size_name'], default)
        return default

    def get_fonts(self):
        """获取全局字体设置"""
        return self.config.get('fonts', {})

    def get_document_settings(self):
        """获取文档级设置"""
        return self.config.get('document', {})

    def get_page_number_config(self, section_name):
        """获取指定部分的页码配置"""
        section_map = {
            'abstract': self.config.get('abstract', {}),
            'toc': self.config.get('toc', {}),
            'body': self.config.get('body', {})
        }
        return section_map.get(section_name, {}).get('page_number', {})

    def apply_paragraph_style(self, paragraph, style_config):
        """
        应用段落样式
        :param paragraph: python-docx的段落对象
        :param style_config: 样式配置字典
        """
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if 'alignment' in style_config:
            paragraph.alignment = alignment_map.get(style_config['alignment'], WD_ALIGN_PARAGRAPH.LEFT)

        paragraph.paragraph_format.space_before = Pt(style_config.get('space_before', 0))
        paragraph.paragraph_format.space_after = Pt(style_config.get('space_after', 0))

        if 'hanging_indent_chars' in style_config:
            char_count = style_config['hanging_indent_chars']
            font_size = self._resolve_size_value(style_config, 12)
            indent_value = Pt(char_count * font_size)
            paragraph.paragraph_format.left_indent = indent_value
            paragraph.paragraph_format.first_line_indent = Pt(-char_count * font_size)
        elif 'first_line_indent' in style_config:
            char_count = style_config['first_line_indent']
            font_size = self._resolve_size_value(style_config, 12)
            indent_twips = int(char_count * font_size * 20)
            paragraph.paragraph_format.first_line_indent = Pt(char_count * font_size)
            self._apply_character_indent(paragraph, char_count, indent_twips)

        if 'line_spacing' in style_config:
            self._apply_line_spacing(paragraph, style_config['line_spacing'])

    def _apply_character_indent(self, paragraph, char_count, indent_twips):
        """通过 XML 设置字符单位的首行缩进"""
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        ind.set(qn('w:firstLine'), str(indent_twips))
        ind.set(qn('w:firstLineChars'), str(int(char_count * 100)))

    def _apply_line_spacing(self, paragraph, spacing_value):
        pf = paragraph.paragraph_format
        if isinstance(spacing_value, (int, float)):
            rule_map = {
                1.0: WD_LINE_SPACING.SINGLE,
                1.5: WD_LINE_SPACING.ONE_POINT_FIVE,
                2.0: WD_LINE_SPACING.DOUBLE
            }
            if spacing_value in rule_map:
                pf.line_spacing_rule = rule_map[spacing_value]
            else:
                pf.line_spacing = spacing_value
            return

        if isinstance(spacing_value, str):
            value = spacing_value.strip().lower()
            if value == 'single':
                pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                return
            if value in {'1.5', 'one_point_five', 'onepointfive'}:
                pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                return
            if value == 'double':
                pf.line_spacing_rule = WD_LINE_SPACING.DOUBLE
                return
            if value.startswith('fixed_') and value.endswith('pt'):
                number_part = value[len('fixed_'):-2]
                try:
                    point_size = float(number_part)
                except ValueError:
                    point_size = 20.0
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = Pt(point_size)
                return
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def apply_run_style(self, run, style_config, text_type='chinese'):
        """
        应用文字样式（Run级别）
        :param run: python-docx的Run对象
        :param style_config: 样式配置字典
        :param text_type: 文本类型（'chinese' 或 'english'）
        """
        r_pr = run._element.get_or_add_rPr()
        if 'font' in style_config:
            font_name = style_config['font']
            run.font.name = font_name
            self._apply_font_family(r_pr, font_name)

        if 'size' in style_config or 'size_name' in style_config:
            run.font.size = Pt(self._resolve_size_value(style_config, style_config.get('size', 12)))

        if 'bold' in style_config:
            self._set_bool_property(r_pr, 'b', style_config['bold'])
            run.font.bold = style_config['bold']

        if 'italic' in style_config:
            self._set_bool_property(r_pr, 'i', style_config['italic'])
            run.font.italic = style_config['italic']

    def set_mixed_font(self, run, text, chinese_font, english_font, size, bold=False):
        """
        为包含中英文混合的文本设置不同字体
        """
        run.text = text
        run.font.name = english_font
        r_pr = run._element.get_or_add_rPr()
        self._apply_font_family(r_pr, english_font, chinese_font)
        run.font.size = Pt(size)
        self._set_bool_property(r_pr, 'b', bold)
        run.font.bold = bold

    def get_abstract_title_style(self):
        """获取摘要标题样式"""
        return self.config['abstract']['title']

    def get_abstract_content_style(self):
        """获取摘要正文样式"""
        return self.config['abstract']['content']

    def get_abstract_keywords_style(self):
        """获取摘要关键词样式"""
        return self.config['abstract']['keywords']

    def get_heading_style(self, level):
        """
        获取标题样式
        """
        heading_map = {
            1: 'heading1',
            2: 'heading2',
            3: 'heading3'
        }
        key = heading_map.get(level, 'heading1')
        return self.config['body'][key]

    def get_paragraph_style(self):
        """获取正文段落样式"""
        return self.config['body']['paragraph']

    def get_figure_style(self):
        """获取图片样式"""
        return self.config['figure']

    def get_table_style(self):
        """获取表格样式"""
        return self.config['table']

    def get_formula_style(self):
        """获取公式样式"""
        return self.config['formula']

    def get_references_style(self):
        """获取参考文献样式"""
        return self.config.get('references', {})

    def get_acknowledgement_style(self):
        """获取致谢样式"""
        return self.config.get('acknowledgements', {})

    def get_appendix_style(self):
        """获取附录样式"""
        return self.config.get('appendix', {})

    def get_toc_style(self):
        return self.config.get('toc', {})

    def _apply_font_family(self, r_pr, ascii_font, east_asia_font=None):
        east_font = east_asia_font or ascii_font
        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)
        r_fonts.set(qn('w:ascii'), ascii_font)
        r_fonts.set(qn('w:hAnsi'), ascii_font)
        r_fonts.set(qn('w:eastAsia'), east_font)

    def _set_bool_property(self, r_pr, tag, enabled):
        element = r_pr.find(qn(f'w:{tag}'))
        if element is None:
            element = OxmlElement(f'w:{tag}')
            r_pr.append(element)
        element.set(qn('w:val'), '1' if enabled else '0')
