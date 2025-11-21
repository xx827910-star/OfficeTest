"""
样式管理器 - 负责管理和应用文档样式
"""
import json
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class StyleManager:
    """管理论文格式样式"""

    def __init__(self, config_path):
        """加载格式配置文件"""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

    def get_font_size(self, size_name):
        """
        获取字号对应的磅值
        :param size_name: 字号名称（如"小四"）
        :return: Pt对象
        """
        size = self.config['sizes'].get(size_name, 12)
        return Pt(size)

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
            'abstract_en': self.config.get('abstract_en', {}),
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
            font_size = style_config.get('size', 12)
            indent_value = Pt(char_count * font_size)
            paragraph.paragraph_format.left_indent = indent_value
            paragraph.paragraph_format.first_line_indent = Pt(-char_count * font_size)
        elif 'first_line_indent' in style_config:
            char_count = style_config['first_line_indent']
            font_size = style_config.get('size', 12)
            indent_twips = int(char_count * font_size * 20)
            paragraph.paragraph_format.first_line_indent = Pt(char_count * font_size)
            self._apply_character_indent(paragraph, char_count, indent_twips)

        # 行距：优先精确磅值，再次是倍数
        if 'line_spacing_pt' in style_config:
            paragraph.paragraph_format.line_spacing = Pt(style_config['line_spacing_pt'])
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        elif 'line_spacing' in style_config:
            spacing = style_config['line_spacing']
            rule_map = {
                1.0: WD_LINE_SPACING.SINGLE,
                1.5: WD_LINE_SPACING.ONE_POINT_FIVE,
                2.0: WD_LINE_SPACING.DOUBLE
            }
            rule = rule_map.get(spacing)
            if rule:
                paragraph.paragraph_format.line_spacing_rule = rule
            else:
                paragraph.paragraph_format.line_spacing = spacing

    def _apply_character_indent(self, paragraph, char_count, indent_twips):
        """通过 XML 设置字符单位的首行缩进"""
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        ind.set(qn('w:firstLine'), str(indent_twips))
        ind.set(qn('w:firstLineChars'), str(int(char_count * 100)))

    def apply_run_style(self, run, style_config, text_type='chinese'):
        """
        应用文字样式（Run级别）
        :param run: python-docx的Run对象
        :param style_config: 样式配置字典
        :param text_type: 文本类型（'chinese' 或 'english'）
        """
        if 'font' in style_config:
            run.font.name = style_config['font']
            run._element.rPr.rFonts.set(qn('w:eastAsia'), style_config['font'])

        if 'size' in style_config:
            run.font.size = Pt(style_config['size'])

        if 'bold' in style_config:
            run.font.bold = style_config['bold']

        if 'italic' in style_config:
            run.font.italic = style_config['italic']

    def set_mixed_font(self, run, text, chinese_font, english_font, size, bold=False):
        """
        为包含中英文混合的文本设置不同字体
        """
        run.text = text
        run.font.name = english_font
        run._element.rPr.rFonts.set(qn('w:eastAsia'), chinese_font)
        run.font.size = Pt(size)
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

    def get_abstract_en_title_style(self):
        """获取英文摘要标题样式"""
        return self.config.get('abstract_en', {}).get('title', {})

    def get_abstract_en_content_style(self):
        """获取英文摘要正文样式"""
        return self.config.get('abstract_en', {}).get('content', {})

    def get_abstract_en_keywords_style(self):
        """获取英文摘要关键词样式"""
        return self.config.get('abstract_en', {}).get('keywords', {})

    def get_toc_config(self):
        """获取目录整体配置"""
        return self.config.get('toc', {})

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
