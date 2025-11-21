"""
中国科学技术大学本科毕业论文格式 - 格式化器

功能：将解析后的内容按照配置生成Word文档
特点：
  1. 人类默认功能：目录TOC域、图表SEQ字段、公式OMML、参考文献交叉引用
  2. 学校特定格式：标题格式、三线表、按章编号等
  3. 配置驱动：所有样式从thesis_format.json读取
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class E1Formatter:
    """中国科学技术大学论文格式化器"""

    def __init__(self, style_manager):
        """
        初始化格式化器

        Args:
            style_manager: E1StyleManager实例
        """
        self.style_manager = style_manager
        self.config = style_manager.get_config()
        self.doc = Document()
        self.bookmark_id = 0
        self.reference_targets = {}
        self.reference_backlinks = {}

    def _get_next_bookmark_id(self):
        """生成下一个书签ID"""
        self.bookmark_id += 1
        return self.bookmark_id

    # ========== 辅助方法：书签和超链接 ==========

    def _add_bookmark_to_paragraph(self, paragraph, bookmark_name):
        """
        在段落中添加书签

        Args:
            paragraph: docx.paragraph对象
            bookmark_name: 书签名称
        """
        bookmark_id = self._get_next_bookmark_id()

        bookmark_start = OxmlElement('w:bookmarkStart')
        bookmark_start.set(qn('w:id'), str(bookmark_id))
        bookmark_start.set(qn('w:name'), bookmark_name)

        bookmark_end = OxmlElement('w:bookmarkEnd')
        bookmark_end.set(qn('w:id'), str(bookmark_id))

        p_element = paragraph._element
        if len(p_element) > 0:
            p_element.insert(0, bookmark_start)
            p_element.append(bookmark_end)
        else:
            p_element.append(bookmark_start)
            p_element.append(bookmark_end)

    def _create_standard_hyperlink(self, paragraph, text, bookmark_name, font_size=12):
        """
        创建黑色超链接

        Args:
            paragraph: 目标段落
            text: 显示文本
            bookmark_name: 跳转目标书签
            font_size: 字号
        """
        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('w:anchor'), bookmark_name)
        hyperlink.set(qn('w:history'), '1')

        run_element = OxmlElement('w:r')
        run_props = OxmlElement('w:rPr')

        style_element = OxmlElement('w:rStyle')
        style_element.set(qn('w:val'), 'Hyperlink')
        run_props.append(style_element)

        fonts_element = OxmlElement('w:rFonts')
        fonts_element.set(qn('w:ascii'), english_font)
        fonts_element.set(qn('w:eastAsia'), chinese_font)
        fonts_element.set(qn('w:hAnsi'), english_font)
        run_props.append(fonts_element)

        size_element = OxmlElement('w:sz')
        size_element.set(qn('w:val'), str(int(font_size * 2)))
        run_props.append(size_element)

        color = OxmlElement('w:color')
        color.set(qn('w:val'), '000000')
        run_props.append(color)

        run_element.append(run_props)

        text_element = OxmlElement('w:t')
        text_element.text = text
        run_element.append(text_element)

        hyperlink.append(run_element)
        paragraph._element.append(hyperlink)

    def _add_pageref_field(self, run, bookmark_name):
        """
        添加PAGEREF域

        Args:
            run: docx.run对象
            bookmark_name: 目标书签名称
        """
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = f' PAGEREF {bookmark_name} \\h '

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        r_element = run._r
        r_element.append(fldChar1)
        r_element.append(instrText)
        r_element.append(fldChar2)

    def _add_tab_stop(self, paragraph, position_cm, alignment='right', leader='dot'):
        """
        添加制表位

        Args:
            paragraph: 段落对象
            position_cm: 制表位位置（厘米）
            alignment: 对齐方式
            leader: 前导符类型
        """
        pPr = paragraph._element.get_or_add_pPr()
        tabs = pPr.find(qn('w:tabs'))
        if tabs is None:
            tabs = OxmlElement('w:tabs')
            pPr.append(tabs)

        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), alignment)
        tab.set(qn('w:leader'), leader)
        tab.set(qn('w:pos'), str(int(position_cm * 567)))  # 1cm = 567 twips
        tabs.append(tab)

    # ========== 辅助方法：SEQ字段 ==========

    def _add_seq_field(self, paragraph, seq_type):
        """
        添加简单SEQ字段（全局编号）

        Args:
            paragraph: 目标段落
            seq_type: 序列类型（如'Figure', 'Table'）
        """
        run = paragraph.add_run()
        r = run._r

        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        r.append(fld_char_begin)

        instr_text = OxmlElement('w:instrText')
        instr_text.set(qn('xml:space'), 'preserve')
        instr_text.text = f' SEQ {seq_type} \\* ARABIC '
        r.append(instr_text)

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        r.append(fld_char_end)

    def _add_chapter_based_seq_field(self, paragraph, seq_type, chapter_num):
        """
        添加按章节编号的SEQ字段

        Args:
            paragraph: 目标段落
            seq_type: 序列类型
            chapter_num: 章节号
        """
        run = paragraph.add_run()
        r = run._r

        fld_char_begin = OxmlElement('w:fldChar')
        fld_char_begin.set(qn('w:fldCharType'), 'begin')
        r.append(fld_char_begin)

        instr_text = OxmlElement('w:instrText')
        instr_text.set(qn('xml:space'), 'preserve')
        instr_text.text = f' SEQ {seq_type}_{chapter_num} \\* ARABIC '
        r.append(instr_text)

        fld_char_end = OxmlElement('w:fldChar')
        fld_char_end.set(qn('w:fldCharType'), 'end')
        r.append(fld_char_end)

    # ========== 辅助方法：OMML公式 ==========

    def _apply_math_justification(self, oMathPara, alignment):
        """设置公式对齐方式"""
        oMathParaPr = oMathPara.find(qn('m:oMathParaPr'))
        if oMathParaPr is None:
            oMathParaPr = OxmlElement('m:oMathParaPr')
            oMathPara.insert(0, oMathParaPr)

        jc = oMathParaPr.find(qn('m:jc'))
        if jc is None:
            jc = OxmlElement('m:jc')
            oMathParaPr.append(jc)

        alignment_map = {
            'left': 'left',
            'center': 'center',
            'right': 'right'
        }
        jc.set(qn('m:val'), alignment_map.get(str(alignment).lower(), 'center'))

    def _add_omml_text_run(self, parent, text, font_name, font_size, italic=False):
        """在OMML中添加文本run"""
        r = OxmlElement('m:r')
        rPr = OxmlElement('m:rPr')
        if italic:
            sty = OxmlElement('m:sty')
            sty.set(qn('m:val'), 'i')
            rPr.append(sty)
        r.append(rPr)

        w_rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:eastAsia'), font_name)
        w_rPr.append(rFonts)

        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), font_size)
        w_rPr.append(sz)

        r.append(w_rPr)

        t = OxmlElement('m:t')
        t.text = text
        r.append(t)

        parent.append(r)

    def _build_omml_runs(self, oMath, text, formula_style):
        """构建OMML公式runs"""
        font_name = formula_style.get('font', 'Times New Roman')
        font_size = str(int(formula_style.get('size', 12) * 2))
        functions = {'sin', 'cos', 'tan', 'floor', 'log', 'ln', 'exp', 'max', 'min', 'sqrt'}

        i = 0
        while i < len(text):
            # 处理下标
            if i < len(text) - 2 and text[i].isalnum():
                j = i
                while j < len(text) and text[j].isalnum():
                    j += 1
                if j < len(text) and text[j] == '_':
                    base = text[i:j]
                    k = j + 1
                    while k < len(text) and text[k].isalnum():
                        k += 1
                    sub = text[j + 1:k]

                    sSub = OxmlElement('m:sSub')
                    e = OxmlElement('m:e')
                    is_var = base.isalpha() and base not in functions
                    self._add_omml_text_run(e, base, font_name, font_size, italic=is_var)
                    sSub.append(e)

                    sub_el = OxmlElement('m:sub')
                    self._add_omml_text_run(sub_el, sub, font_name, font_size, italic=sub.isalpha())
                    sSub.append(sub_el)

                    oMath.append(sSub)
                    i = k
                    continue

            # 处理字母
            if text[i].isalpha():
                j = i
                while j < len(text) and text[j].isalpha():
                    j += 1
                word = text[i:j]
                is_var = len(word) == 1 or word not in functions
                self._add_omml_text_run(oMath, word, font_name, font_size, italic=is_var and word not in functions)
                i = j
                continue

            # 其他字符
            self._add_omml_text_run(oMath, text[i], font_name, font_size, italic=False)
            i += 1

    # ========== 辅助方法：参考文献 ==========

    def _prepare_reference_targets(self, references):
        """准备参考文献书签"""
        self.reference_targets = {}
        for idx in range(len(references)):
            bookmark_name = f'_Reference_{idx + 1}'
            self.reference_targets[idx + 1] = {'bookmark': bookmark_name}

    def _add_internal_reference_link(self, paragraph, text, bookmark_name, font_size=12, bold=False, bookmark_name_for_location=None):
        """添加内部引用链接"""
        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('w:anchor'), bookmark_name)
        hyperlink.set(qn('w:history'), '1')

        run_element = OxmlElement('w:r')
        run_props = OxmlElement('w:rPr')

        fonts_element = OxmlElement('w:rFonts')
        fonts_element.set(qn('w:ascii'), english_font)
        fonts_element.set(qn('w:hAnsi'), english_font)
        fonts_element.set(qn('w:eastAsia'), chinese_font)
        run_props.append(fonts_element)

        size_element = OxmlElement('w:sz')
        size_element.set(qn('w:val'), str(int(font_size * 2)))
        run_props.append(size_element)

        if bold:
            b = OxmlElement('w:b')
            b.set(qn('w:val'), '1')
            run_props.append(b)

        color = OxmlElement('w:color')
        color.set(qn('w:val'), '000000')
        run_props.append(color)

        run_element.append(run_props)
        text_element = OxmlElement('w:t')
        text_element.text = text
        run_element.append(text_element)
        hyperlink.append(run_element)

        if bookmark_name_for_location:
            bookmark_id = self._get_next_bookmark_id()
            start = OxmlElement('w:bookmarkStart')
            start.set(qn('w:id'), str(bookmark_id))
            start.set(qn('w:name'), bookmark_name_for_location)
            end = OxmlElement('w:bookmarkEnd')
            end.set(qn('w:id'), str(bookmark_id))
            p = paragraph._element
            p.append(start)
            p.append(hyperlink)
            p.append(end)
        else:
            paragraph._element.append(hyperlink)

    def _add_text_with_citations(self, paragraph, text, font_size=12):
        """添加包含引用的文本"""
        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        citation_pattern = re.compile(r'\[(\d+)\]')
        last_pos = 0

        for match in citation_pattern.finditer(text):
            # 添加引用前的文本
            if match.start() > last_pos:
                before_text = text[last_pos:match.start()]
                run = paragraph.add_run(before_text)
                self.style_manager.set_mixed_font(run, before_text, chinese_font, english_font, font_size)

            # 添加引用链接
            citation_number = int(match.group(1))
            citation_text = match.group(0)
            target = self.reference_targets.get(citation_number)
            if target:
                bookmark_name = None
                if citation_number not in self.reference_backlinks:
                    bookmark_name = f'_Citation_{citation_number}'
                    self.reference_backlinks[citation_number] = bookmark_name

                self._add_internal_reference_link(
                    paragraph,
                    citation_text,
                    target['bookmark'],
                    font_size=font_size,
                    bold=False,
                    bookmark_name_for_location=bookmark_name
                )
            else:
                run = paragraph.add_run(citation_text)
                self.style_manager.set_mixed_font(run, citation_text, chinese_font, english_font, font_size)

            last_pos = match.end()

        # 添加剩余文本
        if last_pos < len(text):
            remaining_text = text[last_pos:]
            run = paragraph.add_run(remaining_text)
            self.style_manager.set_mixed_font(run, remaining_text, chinese_font, english_font, font_size)

    # ========== 页眉页脚设置 ==========

    def _set_header(self, section, header_text, font_name='宋体', font_size=9):
        """设置页眉"""
        section.header.is_linked_to_previous = False
        header = section.header
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.text = header_text
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for run in header_para.runs:
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            run.font.size = Pt(font_size)

    def _set_page_number(self, section, style='arabic', start_from=1):
        """设置页码"""
        section.footer.is_linked_to_previous = False
        footer = section.footer
        footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 设置页码格式
        sectPr = section._sectPr
        pgNumType = sectPr.find(qn('w:pgNumType'))
        if pgNumType is None:
            pgNumType = OxmlElement('w:pgNumType')
            sectPr.append(pgNumType)

        if style == 'roman':
            pgNumType.set(qn('w:fmt'), 'lowerRoman')
        else:
            pgNumType.set(qn('w:fmt'), 'decimal')

        pgNumType.set(qn('w:start'), str(start_from))

        # 添加PAGE域
        run = footer_para.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = ' PAGE '

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

        run.font.name = 'Times New Roman'
        run.font.size = Pt(9)

    # ========== 主要格式化方法 ==========

    def format_document(self, content):
        """
        格式化整个文档

        Args:
            content: 解析后的内容字典

        Returns:
            Document对象
        """
        # 应用Normal样式默认设置
        self.style_manager.apply_normal_style_defaults(self.doc)

        # 应用页面设置
        self.style_manager.apply_page_setup(self.doc.sections[0])

        # 准备参考文献链接
        self._prepare_reference_targets(content.get('references', []))

        # 生成各部分
        self._generate_abstract(content.get('abstract', {}))
        self._generate_abstract_en(content.get('abstract_en', {}))
        self._generate_toc(content.get('chapters', []))
        self._generate_body(content.get('chapters', []))
        self._generate_references(content.get('references', []))
        self._generate_acknowledgements(content.get('acknowledgements', []))
        self._generate_appendix(content.get('appendix', []))

        return self.doc

    def _generate_abstract(self, abstract_data):
        """生成中文摘要"""
        abstract_config = self.style_manager.get_abstract_style()
        if not abstract_config:
            return

        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        # 标题
        title_config = abstract_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', '摘 要'))
        title_run.font.name = title_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', '黑体'))
        title_run.font.size = Pt(title_config.get('size', 16))
        if title_config.get('bold', False):
            title_run.bold = True
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 正文
        content_config = abstract_config.get('content', {})
        for para_text in abstract_data.get('content', []):
            if not para_text.strip():
                continue
            para = self.doc.add_paragraph()
            run = para.add_run(para_text)
            self.style_manager.set_mixed_font(
                run,
                para_text,
                content_config.get('font_chinese', chinese_font),
                content_config.get('font_english', english_font),
                content_config.get('size', 12)
            )
            self.style_manager.apply_paragraph_style(para, content_config)

        # 关键词
        keywords_config = abstract_config.get('keywords', {})
        keywords = abstract_data.get('keywords', [])
        if keywords:
            kw_para = self.doc.add_paragraph()
            label_run = kw_para.add_run(keywords_config.get('label', '关键词：'))
            label_run.font.name = keywords_config.get('label_font', chinese_font)
            label_run._element.rPr.rFonts.set(qn('w:eastAsia'), keywords_config.get('label_font', chinese_font))
            label_run.font.size = Pt(keywords_config.get('label_size', 12))
            if keywords_config.get('label_bold', True):
                b = label_run._element.rPr.find(qn('w:b'))
                if b is None:
                    b = OxmlElement('w:b')
                    label_run._element.rPr.append(b)
                b.set(qn('w:val'), '1')

            content_text = keywords_config.get('separator', '；').join(keywords)
            content_run = kw_para.add_run(content_text)
            self.style_manager.set_mixed_font(
                content_run,
                content_text,
                keywords_config.get('content_font', chinese_font),
                keywords_config.get('content_font', chinese_font),
                keywords_config.get('content_size', 12)
            )

            self.style_manager.apply_paragraph_style(kw_para, keywords_config)

        # 设置页码
        page_num_config = abstract_config.get('page_number', {})
        if page_num_config:
            section = self.doc.sections[-1]
            self._set_page_number(
                section,
                style=page_num_config.get('style', 'roman'),
                start_from=page_num_config.get('start_from', 1)
            )

    def _generate_abstract_en(self, abstract_data):
        """生成英文摘要"""
        abstract_config = self.style_manager.get_abstract_en_style()
        if not abstract_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        fonts = self.style_manager.get_fonts()
        english_font = fonts.get('english', 'Times New Roman')

        # 标题
        title_config = abstract_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', 'ABSTRACT'))
        title_run.font.name = title_config.get('font', english_font)
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', english_font))
        title_run.font.size = Pt(title_config.get('size', 16))
        if title_config.get('bold', True):
            b = title_run._element.rPr.find(qn('w:b'))
            if b is None:
                b = OxmlElement('w:b')
                title_run._element.rPr.append(b)
            b.set(qn('w:val'), '1')
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 正文
        content_config = abstract_config.get('content', {})
        for para_text in abstract_data.get('content', []):
            if not para_text.strip():
                continue
            para = self.doc.add_paragraph()
            run = para.add_run(para_text)
            run.font.name = content_config.get('font', english_font)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), content_config.get('font', english_font))
            run.font.size = Pt(content_config.get('size', 12))
            self.style_manager.apply_paragraph_style(para, content_config)

        # 关键词
        keywords_config = abstract_config.get('keywords', {})
        keywords = abstract_data.get('keywords', [])
        if keywords:
            # 首字母大写
            if keywords_config.get('capitalize_first_letter', True):
                keywords = [kw.capitalize() if kw else kw for kw in keywords]

            kw_para = self.doc.add_paragraph()
            label_run = kw_para.add_run(keywords_config.get('label', 'Key Words:'))
            label_run.font.name = keywords_config.get('label_font', english_font)
            label_run._element.rPr.rFonts.set(qn('w:eastAsia'), keywords_config.get('label_font', english_font))
            label_run.font.size = Pt(keywords_config.get('label_size', 12))
            if keywords_config.get('label_bold', True):
                b = label_run._element.rPr.find(qn('w:b'))
                if b is None:
                    b = OxmlElement('w:b')
                    label_run._element.rPr.append(b)
                b.set(qn('w:val'), '1')

            content_text = keywords_config.get('separator', ';').join(keywords)
            content_run = kw_para.add_run(content_text)
            content_run.font.name = keywords_config.get('content_font', english_font)
            content_run._element.rPr.rFonts.set(qn('w:eastAsia'), keywords_config.get('content_font', english_font))
            content_run.font.size = Pt(keywords_config.get('content_size', 12))

            self.style_manager.apply_paragraph_style(kw_para, keywords_config)

        # 设置页码
        page_num_config = abstract_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(
                section,
                style=page_num_config.get('style', 'roman')
            )

    def _generate_toc(self, chapters):
        """生成目录"""
        toc_config = self.style_manager.get_toc_style()
        if not toc_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')

        # 标题
        title_config = toc_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', '目 录'))
        title_run.font.name = title_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', '黑体'))
        title_run.font.size = Pt(title_config.get('size', 16))
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 目录条目
        entries_config = toc_config.get('entries', {})
        indentation_config = toc_config.get('indentation', {})

        for chapter in chapters:
            # 章标题
            entry_para = self.doc.add_paragraph()
            chapter_num = chapter.get('number', 1)
            chapter_title = chapter.get('title', '')

            # 中文数字转换
            chinese_numbers = self.config.get('chinese_number_mapping', {})
            chinese_num_reverse = {v: k for k, v in chinese_numbers.items()}
            chinese_chapter_num = chinese_num_reverse.get(chapter_num, str(chapter_num))

            full_title = f"第{chinese_chapter_num}章 {chapter_title}"
            bookmark_name = f"_Chapter_{chapter_num}"

            # 使用超链接
            self._create_standard_hyperlink(
                entry_para,
                full_title,
                bookmark_name,
                font_size=entries_config.get('size', 12)
            )

            # 添加制表位和页码
            self._add_tab_stop(entry_para, position_cm=16.0, alignment='right', leader='dot')
            entry_para.add_run('\t')

            page_run = entry_para.add_run()
            self._add_pageref_field(page_run, bookmark_name)
            page_run.font.size = Pt(entries_config.get('size', 12))

            # 加粗章标题
            if indentation_config.get('chapter_bold', True):
                for run in entry_para.runs:
                    r_pr = run._element.get_or_add_rPr()
                    b = r_pr.find(qn('w:b'))
                    if b is None:
                        b = OxmlElement('w:b')
                        r_pr.append(b)
                    b.set(qn('w:val'), '1')

            entry_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            entry_para.paragraph_format.line_spacing = Pt(entries_config.get('line_spacing', 20))

        # 设置页码
        page_num_config = toc_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(
                section,
                style=page_num_config.get('style', 'roman')
            )

    def _generate_body(self, chapters):
        """生成正文"""
        body_config = self.style_manager.get_body_style()
        if not body_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        # 设置页眉
        header_config = body_config.get('header', {})
        if header_config:
            self._set_header(
                section,
                header_config.get('content', '论文题目'),
                font_name=header_config.get('font', '宋体'),
                font_size=header_config.get('size', 9)
            )

        # 设置页码
        page_num_config = body_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(
                section,
                style=page_num_config.get('style', 'arabic'),
                start_from=page_num_config.get('start_from', 1)
            )

        # 处理每一章
        for chapter in chapters:
            self._generate_chapter(chapter)

    def _generate_chapter(self, chapter):
        """生成章节内容"""
        chapter_num = chapter.get('number', 1)
        chapter_title = chapter.get('title', '')

        # 生成一级标题
        self._add_heading1(chapter_num, chapter_title)

        # 生成章节内容
        for item in chapter.get('content', []):
            item_type = item.get('type')

            if item_type == 'heading2':
                self._add_heading2(item)
            elif item_type == 'heading3':
                self._add_heading3(item)
            elif item_type == 'heading4':
                self._add_heading4(item)
            elif item_type == 'heading5':
                self._add_heading5(item)
            elif item_type == 'paragraph':
                self._add_paragraph(item, chapter_num)
            elif item_type == 'figure':
                self._add_figure(item, chapter_num)
            elif item_type == 'table':
                self._add_table(item, chapter_num)
            elif item_type == 'formula':
                self._add_formula(item)

    def _add_heading1(self, chapter_num, title):
        """添加一级标题"""
        body_config = self.style_manager.get_body_style()
        h1_config = body_config.get('heading1', {})

        # 中文数字转换
        chinese_numbers = self.config.get('chinese_number_mapping', {})
        chinese_num_reverse = {v: k for k, v in chinese_numbers.items()}
        chinese_chapter_num = chinese_num_reverse.get(chapter_num, str(chapter_num))

        para = self.doc.add_paragraph()
        bookmark_name = f"_Chapter_{chapter_num}"
        self._add_bookmark_to_paragraph(para, bookmark_name)

        # 编号
        number_format = h1_config.get('format', '第{chinese}章')
        number_text = number_format.replace('{chinese}', chinese_chapter_num)
        number_run = para.add_run(number_text)
        number_run.font.name = h1_config.get('number_font', '黑体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), h1_config.get('number_font', '黑体'))
        number_run.font.size = Pt(h1_config.get('size', 16))

        # 标题
        para.add_run(' ')
        title_run = para.add_run(title)
        title_run.font.name = h1_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), h1_config.get('font', '黑体'))
        title_run.font.size = Pt(h1_config.get('size', 16))

        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(h1_config.get('space_before', 24))
        para.paragraph_format.space_after = Pt(h1_config.get('space_after', 18))
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def _add_heading2(self, item):
        """添加二级标题"""
        body_config = self.style_manager.get_body_style()
        h2_config = body_config.get('heading2', {})

        chinese_section_num = item.get('chinese_number', '一')

        para = self.doc.add_paragraph()

        # 编号
        number_format = h2_config.get('format', '第{chinese}节')
        number_text = number_format.replace('{chinese}', chinese_section_num)
        number_run = para.add_run(number_text)
        number_run.font.name = h2_config.get('number_font', '黑体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), h2_config.get('number_font', '黑体'))
        number_run.font.size = Pt(h2_config.get('size', 15))

        # 标题
        para.add_run(' ')
        title_run = para.add_run(item.get('text', ''))
        title_run.font.name = h2_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), h2_config.get('font', '黑体'))
        title_run.font.size = Pt(h2_config.get('size', 15))

        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(h2_config.get('space_before', 18))
        para.paragraph_format.space_after = Pt(h2_config.get('space_after', 12))
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def _add_heading3(self, item):
        """添加三级标题"""
        body_config = self.style_manager.get_body_style()
        h3_config = body_config.get('heading3', {})

        chinese_subsection_num = item.get('chinese_number', '一')

        para = self.doc.add_paragraph()

        # 编号
        number_format = h3_config.get('format', '{chinese}、')
        number_text = number_format.replace('{chinese}', chinese_subsection_num)
        number_run = para.add_run(number_text)
        number_run.font.name = h3_config.get('number_font', '黑体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), h3_config.get('number_font', '黑体'))
        number_run.font.size = Pt(h3_config.get('size', 14))

        # 标题
        if item.get('text'):
            title_run = para.add_run(item.get('text', ''))
            title_run.font.name = h3_config.get('font', '黑体')
            title_run._element.rPr.rFonts.set(qn('w:eastAsia'), h3_config.get('font', '黑体'))
            title_run.font.size = Pt(h3_config.get('size', 14))

        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(h3_config.get('space_before', 12))
        para.paragraph_format.space_after = Pt(h3_config.get('space_after', 6))
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    def _add_heading4(self, item):
        """添加四级标题"""
        body_config = self.style_manager.get_body_style()
        h4_config = body_config.get('heading4', {})

        para = self.doc.add_paragraph()

        # 编号
        number_format = h4_config.get('format', '{arabic}.')
        number_text = number_format.replace('{arabic}', str(item.get('number', '1')))
        number_run = para.add_run(number_text)
        number_run.font.name = h4_config.get('font', '宋体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), h4_config.get('font', '宋体'))
        number_run.font.size = Pt(h4_config.get('size', 12))

        # 标题
        para.add_run(' ')
        title_run = para.add_run(item.get('text', ''))
        fonts = self.style_manager.get_fonts()
        self.style_manager.set_mixed_font(
            title_run,
            item.get('text', ''),
            h4_config.get('font', fonts.get('chinese', '宋体')),
            fonts.get('english', 'Times New Roman'),
            h4_config.get('size', 12)
        )

        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(h4_config.get('space_before', 6))
        para.paragraph_format.space_after = Pt(h4_config.get('space_after', 3))
        if h4_config.get('line_spacing_rule') == 'fixed':
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            para.paragraph_format.line_spacing = Pt(h4_config.get('line_spacing', 20))

    def _add_heading5(self, item):
        """添加五级标题"""
        body_config = self.style_manager.get_body_style()
        h5_config = body_config.get('heading5', {})

        para = self.doc.add_paragraph()

        # 编号
        number_format = h5_config.get('format', '（{arabic}）')
        number_text = number_format.replace('{arabic}', str(item.get('number', '1')))
        number_run = para.add_run(number_text)
        number_run.font.name = h5_config.get('font', '宋体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), h5_config.get('font', '宋体'))
        number_run.font.size = Pt(h5_config.get('size', 12))

        # 标题
        para.add_run(' ')
        title_run = para.add_run(item.get('text', ''))
        fonts = self.style_manager.get_fonts()
        self.style_manager.set_mixed_font(
            title_run,
            item.get('text', ''),
            h5_config.get('font', fonts.get('chinese', '宋体')),
            fonts.get('english', 'Times New Roman'),
            h5_config.get('size', 12)
        )

        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(h5_config.get('space_before', 3))
        para.paragraph_format.space_after = Pt(h5_config.get('space_after', 3))
        if h5_config.get('line_spacing_rule') == 'fixed':
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            para.paragraph_format.line_spacing = Pt(h5_config.get('line_spacing', 20))

    def _add_paragraph(self, item, chapter_num):
        """添加段落"""
        text = item.get('text', '')
        if not text.strip():
            return

        para = self.doc.add_paragraph()
        paragraph_config = self.style_manager.get_paragraph_style()

        # 处理引用
        self._add_text_with_citations(para, text, font_size=paragraph_config.get('size', 12))

        # 应用段落样式
        self.style_manager.apply_paragraph_style(para, paragraph_config)

    def _add_figure(self, item, chapter_num):
        """添加图片"""
        figure_config = self.style_manager.get_figure_style()
        image_path = item.get('path')

        # 图片段落
        img_para = self.doc.add_paragraph()
        img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if not image_path or not os.path.exists(image_path):
            # 占位符
            placeholder_run = img_para.add_run('[图片缺失: 请检查路径]')
            placeholder_run.font.color.rgb = RGBColor(255, 0, 0)
            placeholder_run.font.size = Pt(12)
        else:
            try:
                run = img_para.add_run()
                run.add_picture(image_path, width=Inches(figure_config.get('width_in', 5)))
            except Exception as e:
                error_run = img_para.add_run(f'[图片加载失败: {str(e)}]')
                error_run.font.color.rgb = RGBColor(255, 0, 0)
                error_run.font.size = Pt(12)

        # 题注段落
        caption_para = self.doc.add_paragraph()
        caption_config = figure_config.get('caption', {})

        # 使用SEQ字段自动编号
        label_run = caption_para.add_run('图')
        label_run.font.name = caption_config.get('font', '宋体')
        label_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        label_run.font.size = Pt(caption_config.get('size', 12))

        num_run = caption_para.add_run(str(chapter_num))
        num_run.font.name = caption_config.get('font', '宋体')
        num_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        num_run.font.size = Pt(caption_config.get('size', 12))

        sep_run = caption_para.add_run('-')
        sep_run.font.name = caption_config.get('font', '宋体')
        sep_run.font.size = Pt(caption_config.get('size', 12))

        self._add_chapter_based_seq_field(caption_para, 'Figure', chapter_num)

        # 题注文字
        caption_text = caption_config.get('label_separator', ' ') + item.get('caption', '')
        caption_run = caption_para.add_run(caption_text)
        caption_run.font.name = caption_config.get('font', '宋体')
        caption_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        caption_run.font.size = Pt(caption_config.get('size', 12))

        caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_para.paragraph_format.space_before = Pt(caption_config.get('space_before', 6))
        caption_para.paragraph_format.space_after = Pt(caption_config.get('space_after', 6))

        # 来源
        if item.get('source'):
            source_para = self.doc.add_paragraph()
            source_config = figure_config.get('source', {})
            source_run = source_para.add_run(f"来源：{item['source']}")
            source_run.font.name = source_config.get('font', '宋体')
            source_run._element.rPr.rFonts.set(qn('w:eastAsia'), source_config.get('font', '宋体'))
            source_run.font.size = Pt(source_config.get('size', 9))
            source_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _add_table(self, item, chapter_num):
        """添加表格"""
        table_config = self.style_manager.get_table_style()
        table_data = item.get('data', [])
        if not table_data:
            return

        # 题注段落
        caption_para = self.doc.add_paragraph()
        caption_config = table_config.get('caption', {})

        # 使用SEQ字段自动编号
        label_run = caption_para.add_run('表')
        label_run.font.name = caption_config.get('font', '宋体')
        label_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        label_run.font.size = Pt(caption_config.get('size', 12))

        num_run = caption_para.add_run(str(chapter_num))
        num_run.font.name = caption_config.get('font', '宋体')
        num_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        num_run.font.size = Pt(caption_config.get('size', 12))

        sep_run = caption_para.add_run('-')
        sep_run.font.name = caption_config.get('font', '宋体')
        sep_run.font.size = Pt(caption_config.get('size', 12))

        self._add_chapter_based_seq_field(caption_para, 'Table', chapter_num)

        # 题注文字
        caption_text = caption_config.get('label_separator', ' ') + item.get('caption', '')
        caption_run = caption_para.add_run(caption_text)
        caption_run.font.name = caption_config.get('font', '宋体')
        caption_run._element.rPr.rFonts.set(qn('w:eastAsia'), caption_config.get('font', '宋体'))
        caption_run.font.size = Pt(caption_config.get('size', 12))

        caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_para.paragraph_format.space_before = Pt(caption_config.get('space_before', 6))
        caption_para.paragraph_format.space_after = Pt(caption_config.get('space_after', 6))

        # 创建表格
        rows = len(table_data)
        cols = max(len(row) for row in table_data) if table_data else 0
        table = self.doc.add_table(rows=rows, cols=cols)

        # 填充表格内容
        for i, row_data in enumerate(table_data):
            row = table.rows[i]
            for j, cell_text in enumerate(row_data):
                if j < len(row.cells):
                    cell = row.cells[j]
                    cell.text = cell_text
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in paragraph.runs:
                            run.font.name = table_config.get('content_font', '宋体')
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), table_config.get('content_font', '宋体'))
                            run.font.size = Pt(table_config.get('content_size', 12))

        # 设置三线表边框
        if table_config.get('border_style') == 'three_line':
            self._set_three_line_borders(table, table_config)

        # 来源
        if item.get('source'):
            source_para = self.doc.add_paragraph()
            source_config = table_config.get('source', {})
            source_run = source_para.add_run(f"来源：{item['source']}")
            source_run.font.name = source_config.get('font', '宋体')
            source_run._element.rPr.rFonts.set(qn('w:eastAsia'), source_config.get('font', '宋体'))
            source_run.font.size = Pt(source_config.get('size', 9))
            source_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def _set_three_line_borders(self, table, table_config):
        """设置三线表边框"""
        tbl = table._element
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)

        # 移除默认边框
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)

        # 设置上边框
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), str(int(table_config.get('top_border', 1.5) * 8)))
        top.set(qn('w:space'), '0')
        top.set(qn('w:color'), '000000')
        tblBorders.append(top)

        # 设置下边框
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), str(int(table_config.get('bottom_border', 1.5) * 8)))
        bottom.set(qn('w:space'), '0')
        bottom.set(qn('w:color'), '000000')
        tblBorders.append(bottom)

        # 移除左右边框
        for border_name in ['left', 'right', 'insideV']:
            border_elem = OxmlElement(f'w:{border_name}')
            border_elem.set(qn('w:val'), 'none')
            tblBorders.append(border_elem)

        # 设置中线（第一行底部）
        if len(table.rows) > 1:
            first_row = table.rows[0]
            tr = first_row._tr
            trPr = tr.get_or_add_trPr()

            for cell in first_row.cells:
                tcPr = cell._element.get_or_add_tcPr()
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is None:
                    tcBorders = OxmlElement('w:tcBorders')
                    tcPr.append(tcBorders)

                bottom_border = OxmlElement('w:bottom')
                bottom_border.set(qn('w:val'), 'single')
                bottom_border.set(qn('w:sz'), str(int(table_config.get('middle_border', 0.5) * 8)))
                bottom_border.set(qn('w:space'), '0')
                bottom_border.set(qn('w:color'), '000000')
                tcBorders.append(bottom_border)

    def _add_formula(self, item):
        """添加公式"""
        formula_config = self.style_manager.get_formula_style()
        formula_content = item.get('content', '')
        if not formula_content.strip():
            return

        # 公式段落
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 设置制表位（公式居中，编号右对齐）
        self._add_tab_stop(p, position_cm=8.25, alignment='center', leader='none')
        self._add_tab_stop(p, position_cm=16.0, alignment='right', leader='none')

        p.add_run('\t')

        # 构建OMML公式
        try:
            oMathPara = OxmlElement('m:oMathPara')
            oMath = OxmlElement('m:oMath')
            self._build_omml_runs(oMath, formula_content, formula_config)
            oMathPara.append(oMath)
            self._apply_math_justification(oMathPara, formula_config.get('alignment', 'center'))
            p._element.append(oMathPara)
        except Exception as e:
            print(f"OMML构建失败: {str(e)}，使用文本模式")
            run = p.add_run(formula_content)
            run.font.name = formula_config.get('font', 'Times New Roman')
            run.font.size = Pt(formula_config.get('size', 12))

        # 公式编号
        p.add_run('\t')
        number_format = formula_config.get('number_format', '（{seq}）')
        # 简化：这里直接显示编号，实际应该用SEQ
        number_text = number_format.replace('{seq}', item.get('number', '1'))
        number_run = p.add_run(number_text)
        number_run.font.name = formula_config.get('number_font', '宋体')
        number_run._element.rPr.rFonts.set(qn('w:eastAsia'), formula_config.get('number_font', '宋体'))
        number_run.font.size = Pt(formula_config.get('number_size', 12))

        if formula_config.get('number_line_spacing_rule') == 'fixed':
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(formula_config.get('number_line_spacing', 20))

    def _generate_references(self, references):
        """生成参考文献"""
        if not references:
            return

        references_config = self.style_manager.get_references_style()
        if not references_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        # 设置页眉
        header_config = references_config.get('header', {})
        if header_config:
            self._set_header(
                section,
                header_config.get('content', '参考文献'),
                font_name=header_config.get('font', '宋体'),
                font_size=header_config.get('size', 9)
            )

        # 设置页码
        page_num_config = references_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(section, style=page_num_config.get('style', 'arabic'))

        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        # 标题
        title_config = references_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', '参 考 文 献'))
        title_run.font.name = title_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', '黑体'))
        title_run.font.size = Pt(title_config.get('size', 16))
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 文献条目
        entry_config = references_config.get('entry', {})
        number_config = references_config.get('number', {})

        for idx, ref_text in enumerate(references):
            para = self.doc.add_paragraph()
            bookmark_name = f'_Reference_{idx + 1}'
            self._add_bookmark_to_paragraph(para, bookmark_name)

            # 编号
            number_format = number_config.get('format', '{index}.')
            number_text = number_format.replace('{index}', str(idx + 1))
            number_run = para.add_run(number_text)
            number_run.font.name = number_config.get('font', english_font)
            number_run.font.size = Pt(entry_config.get('size', 12))

            # 文献内容
            content_run = para.add_run(' ' + ref_text)
            self.style_manager.set_mixed_font(
                content_run,
                ' ' + ref_text,
                entry_config.get('font_chinese', chinese_font),
                entry_config.get('font_english', english_font),
                entry_config.get('size', 12)
            )

            self.style_manager.apply_paragraph_style(para, entry_config)

    def _generate_acknowledgements(self, acknowledgements):
        """生成致谢"""
        if not acknowledgements:
            return

        ack_config = self.style_manager.get_acknowledgements_style()
        if not ack_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        # 设置页眉
        header_config = ack_config.get('header', {})
        if header_config:
            self._set_header(
                section,
                header_config.get('content', '致谢'),
                font_name=header_config.get('font', '宋体'),
                font_size=header_config.get('size', 9)
            )

        # 设置页码
        page_num_config = ack_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(section, style=page_num_config.get('style', 'arabic'))

        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        # 标题
        title_config = ack_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', '致 谢'))
        title_run.font.name = title_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', '黑体'))
        title_run.font.size = Pt(title_config.get('size', 16))
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 正文
        content_config = ack_config.get('content', {})
        for para_text in acknowledgements:
            if not para_text.strip():
                continue
            para = self.doc.add_paragraph()
            run = para.add_run(para_text)
            self.style_manager.set_mixed_font(
                run,
                para_text,
                content_config.get('font_chinese', chinese_font),
                content_config.get('font_english', english_font),
                content_config.get('size', 12)
            )
            self.style_manager.apply_paragraph_style(para, content_config)

    def _generate_appendix(self, appendix):
        """生成附录"""
        if not appendix:
            return

        app_config = self.style_manager.get_appendix_style()
        if not app_config:
            return

        # 新建分节
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        section = self.doc.sections[-1]
        self.style_manager.apply_page_setup(section)

        # 设置页眉
        header_config = app_config.get('header', {})
        if header_config:
            self._set_header(
                section,
                header_config.get('content', '附录'),
                font_name=header_config.get('font', '宋体'),
                font_size=header_config.get('size', 9)
            )

        # 设置页码
        page_num_config = app_config.get('page_number', {})
        if page_num_config:
            self._set_page_number(section, style=page_num_config.get('style', 'arabic'))

        fonts = self.style_manager.get_fonts()
        chinese_font = fonts.get('chinese', '宋体')
        english_font = fonts.get('english', 'Times New Roman')

        # 标题
        title_config = app_config.get('title', {})
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_config.get('text', '附 录'))
        title_run.font.name = title_config.get('font', '黑体')
        title_run._element.rPr.rFonts.set(qn('w:eastAsia'), title_config.get('font', '黑体'))
        title_run.font.size = Pt(title_config.get('size', 16))
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_para.paragraph_format.space_before = Pt(title_config.get('space_before', 24))
        title_para.paragraph_format.space_after = Pt(title_config.get('space_after', 18))
        title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        # 正文
        content_config = app_config.get('content', {})
        for para_text in appendix:
            if not para_text.strip():
                continue
            para = self.doc.add_paragraph()
            run = para.add_run(para_text)
            self.style_manager.set_mixed_font(
                run,
                para_text,
                content_config.get('font_chinese', chinese_font),
                content_config.get('font_english', english_font),
                content_config.get('size', 12)
            )
            self.style_manager.apply_paragraph_style(para, content_config)
