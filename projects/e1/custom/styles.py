"""
中国科学技术大学本科毕业论文格式 - 样式管理器

功能：从配置文件读取并提供格式配置，应用段落和run级别的样式
特点：配置驱动，所有样式参数从thesis_format.json读取
"""

import json
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class E1StyleManager:
    """中国科学技术大学论文格式样式管理器"""

    def __init__(self, config_path):
        """
        初始化样式管理器

        Args:
            config_path: thesis_format.json配置文件路径
        """
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

    def get_config(self):
        """获取完整配置"""
        return self.config

    def get_fonts(self):
        """获取字体配置"""
        return self.config.get('fonts', {})

    def get_abstract_style(self):
        """获取中文摘要样式"""
        return self.config.get('abstract', {})

    def get_abstract_en_style(self):
        """获取英文摘要样式"""
        return self.config.get('abstract_en', {})

    def get_toc_style(self):
        """获取目录样式"""
        return self.config.get('toc', {})

    def get_body_style(self):
        """获取正文样式"""
        return self.config.get('body', {})

    def get_paragraph_style(self):
        """获取正文段落样式"""
        return self.config.get('body', {}).get('paragraph', {})

    def get_figure_style(self):
        """获取图片样式"""
        return self.config.get('figure', {})

    def get_table_style(self):
        """获取表格样式"""
        return self.config.get('table', {})

    def get_formula_style(self):
        """获取公式样式"""
        return self.config.get('formula', {})

    def get_references_style(self):
        """获取参考文献样式"""
        return self.config.get('references', {})

    def get_acknowledgements_style(self):
        """获取致谢样式"""
        return self.config.get('acknowledgements', {})

    def get_appendix_style(self):
        """获取附录样式"""
        return self.config.get('appendix', {})

    def set_mixed_font(self, run, text, chinese_font, english_font, size, bold=False, italic=False):
        """
        设置中英文混排字体

        Args:
            run: docx.text.run.Run对象
            text: 文本内容
            chinese_font: 中文字体名称
            english_font: 英文字体名称
            size: 字号（磅）
            bold: 是否加粗
            italic: 是否斜体
        """
        run.text = text
        run.font.name = english_font
        run.font.size = Pt(size)

        # 设置XML级别的字体属性
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)

        r_fonts.set(qn('w:ascii'), english_font)
        r_fonts.set(qn('w:eastAsia'), chinese_font)
        r_fonts.set(qn('w:hAnsi'), english_font)

        # 设置加粗
        if bold:
            b = r_pr.find(qn('w:b'))
            if b is None:
                b = OxmlElement('w:b')
                r_pr.append(b)
            b.set(qn('w:val'), '1')

        # 设置斜体
        if italic:
            i = r_pr.find(qn('w:i'))
            if i is None:
                i = OxmlElement('w:i')
                r_pr.append(i)
            i.set(qn('w:val'), '1')

        # 设置字号（XML级别）
        sz = r_pr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            r_pr.append(sz)
        sz.set(qn('w:val'), str(int(size * 2)))

        sz_cs = r_pr.find(qn('w:szCs'))
        if sz_cs is None:
            sz_cs = OxmlElement('w:szCs')
            r_pr.append(sz_cs)
        sz_cs.set(qn('w:val'), str(int(size * 2)))

    def apply_run_style(self, run, style_config, text_type='mixed'):
        """
        应用run级别样式

        Args:
            run: docx.text.run.Run对象
            style_config: 样式配置字典
            text_type: 文本类型('chinese', 'english', 'mixed')
        """
        # 设置字体
        if 'font' in style_config:
            run.font.name = style_config['font']
            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.find(qn('w:rFonts'))
            if r_fonts is None:
                r_fonts = OxmlElement('w:rFonts')
                r_pr.append(r_fonts)
            r_fonts.set(qn('w:ascii'), style_config['font'])
            r_fonts.set(qn('w:eastAsia'), style_config['font'])
            r_fonts.set(qn('w:hAnsi'), style_config['font'])

        # 设置字号
        if 'size' in style_config:
            run.font.size = Pt(style_config['size'])
            r_pr = run._element.get_or_add_rPr()
            sz = r_pr.find(qn('w:sz'))
            if sz is None:
                sz = OxmlElement('w:sz')
                r_pr.append(sz)
            sz.set(qn('w:val'), str(int(style_config['size'] * 2)))

        # 设置加粗
        if 'bold' in style_config:
            r_pr = run._element.get_or_add_rPr()
            b = r_pr.find(qn('w:b'))
            if b is None:
                b = OxmlElement('w:b')
                r_pr.append(b)
            b.set(qn('w:val'), '1' if style_config['bold'] else '0')

        # 设置斜体
        if 'italic' in style_config:
            r_pr = run._element.get_or_add_rPr()
            i = r_pr.find(qn('w:i'))
            if i is None:
                i = OxmlElement('w:i')
                r_pr.append(i)
            i.set(qn('w:val'), '1' if style_config['italic'] else '0')

    def apply_paragraph_style(self, paragraph, style_config):
        """
        应用段落级别样式

        Args:
            paragraph: docx.text.paragraph.Paragraph对象
            style_config: 样式配置字典
        """
        # 对齐方式
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if 'alignment' in style_config:
            paragraph.alignment = alignment_map.get(
                style_config['alignment'],
                WD_ALIGN_PARAGRAPH.LEFT
            )

        # 段前段后间距
        paragraph.paragraph_format.space_before = Pt(style_config.get('space_before', 0))
        paragraph.paragraph_format.space_after = Pt(style_config.get('space_after', 0))

        # 首行缩进
        if 'first_line_indent' in style_config:
            char_count = style_config['first_line_indent']
            font_size = style_config.get('size', 12)
            paragraph.paragraph_format.first_line_indent = Pt(char_count * font_size)

        # 左缩进（用于悬挂缩进）
        if 'left_indent_pt' in style_config:
            paragraph.paragraph_format.left_indent = Pt(style_config['left_indent_pt'])

        # 首行悬挂缩进
        if 'first_line_indent_pt' in style_config:
            paragraph.paragraph_format.first_line_indent = Pt(style_config['first_line_indent_pt'])

        # 悬挂缩进（字符数）
        if 'hanging_indent_chars' in style_config:
            char_count = style_config['hanging_indent_chars']
            font_size = style_config.get('size', 12)
            paragraph.paragraph_format.left_indent = Pt(char_count * font_size)
            paragraph.paragraph_format.first_line_indent = Pt(-char_count * font_size)

        # 行距
        if 'line_spacing_rule' in style_config:
            if style_config['line_spacing_rule'] == 'fixed':
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                if 'line_spacing' in style_config:
                    paragraph.paragraph_format.line_spacing = Pt(style_config['line_spacing'])
            elif style_config['line_spacing_rule'] == 'multiple':
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                if 'line_spacing' in style_config:
                    paragraph.paragraph_format.line_spacing = style_config['line_spacing']
        elif 'line_spacing' in style_config:
            spacing = style_config['line_spacing']
            if spacing == 1.0:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            elif spacing == 1.5:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            elif spacing == 2.0:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
            else:
                paragraph.paragraph_format.line_spacing = spacing

    def apply_normal_style_defaults(self, document):
        """
        应用Normal样式的默认设置

        Args:
            document: docx.Document对象
        """
        fonts = self.get_fonts()
        paragraph_style = self.get_paragraph_style()

        english_font = fonts.get('english', 'Times New Roman')
        chinese_font = fonts.get('chinese', '宋体')
        font_size = paragraph_style.get('size', 12)

        # 设置Normal样式
        normal_style = document.styles['Normal']
        font = normal_style.font
        font.name = english_font
        font.size = Pt(font_size)

        # 设置东亚字体
        rPr = normal_style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), english_font)
        rFonts.set(qn('w:eastAsia'), chinese_font)
        rFonts.set(qn('w:hAnsi'), english_font)

        # 设置段落格式
        pPr = normal_style.element.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        spacing.set(qn('w:line'), str(int(20 * 20)))  # 固定20磅行距
        spacing.set(qn('w:lineRule'), 'exact')

    def get_page_setup(self):
        """获取页面设置"""
        return self.config.get('document', {})

    def apply_page_setup(self, section):
        """
        应用页面设置

        Args:
            section: docx.section.Section对象
        """
        page_setup = self.get_page_setup()

        # 设置页边距
        margins = page_setup.get('margins', {})
        section.top_margin = Cm(margins.get('top', 2.5))
        section.bottom_margin = Cm(margins.get('bottom', 2.5))
        section.left_margin = Cm(margins.get('left', 2.5))
        section.right_margin = Cm(margins.get('right', 2.5))

        # 设置页眉页脚距离
        section.header_distance = Cm(page_setup.get('header_distance', 1.4))
        section.footer_distance = Cm(page_setup.get('footer_distance', 1.3))

        # 设置装订线
        section.gutter = Cm(page_setup.get('gutter', 0))
