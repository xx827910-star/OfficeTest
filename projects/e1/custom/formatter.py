import os
import re
from typing import Any, Dict, List, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from src.generator import ThesisGenerator


class E1ThesisFormatter(ThesisGenerator):
    """
    Formatter with e1-specific customisations and audit logging.
    """

    def _reset_document(self):
        self.audit_log = {
            'toc_inserted': False,
            'bookmark_count': 0,
            'seq_fields_used': False,
            'heading_runs_have_east_asia': False,
            'paragraph_spacing_applied': False,
            'missing_figures': [],
            'omml_used': False,
        }
        super()._reset_document()
        self._enable_field_updates()

    def _apply_math_defaults(self):
        super()._apply_math_defaults()
        if hasattr(self, 'audit_log'):
            self.audit_log['omml_used'] = True

    # ------------------------------------------------------------------
    # Overrides
    # ------------------------------------------------------------------
    def _add_bookmark_to_paragraph(self, paragraph, bookmark_name):
        super()._add_bookmark_to_paragraph(paragraph, bookmark_name)
        self.audit_log['bookmark_count'] += 1

    def _generate_abstract(self, abstract_data):
        config = self.style_manager.config.get('abstract', {})
        title_cfg = config.get('title', {})
        title_text = title_cfg.get('text', '摘 要')
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_text)
        self.style_manager.set_mixed_font(
            title_run,
            title_text,
            chinese_font=title_cfg.get('font', self.style_manager.get_fonts().get('chinese')),
            english_font=title_cfg.get('font', self.style_manager.get_fonts().get('english')),
            size=title_cfg.get('size'),
            bold=title_cfg.get('bold', False),
        )
        self.style_manager.apply_paragraph_style(title_para, title_cfg)

        for para_text in abstract_data.get('content', []):
            para = self.doc.add_paragraph()
            body_cfg = config.get('content', {})
            run = para.add_run()
            self.style_manager.set_mixed_font(
                run,
                para_text,
                chinese_font=body_cfg.get('font'),
                english_font=self.style_manager.get_fonts().get('english'),
                size=body_cfg.get('size'),
            )
            self.style_manager.apply_paragraph_style(para, body_cfg)

        keywords = abstract_data.get('keywords') or []
        if keywords:
            kw_cfg = config.get('keywords', {})
            kw_para = self.doc.add_paragraph()
            label_text = kw_cfg.get('label_text', '关键词：')
            label_run = kw_para.add_run(label_text)
            self.style_manager.set_mixed_font(
                label_run,
                label_text,
                chinese_font=kw_cfg.get('label_font'),
                english_font=self.style_manager.get_fonts().get('english'),
                size=kw_cfg.get('label_size'),
                bold=kw_cfg.get('label_bold', True),
            )
            separator = kw_cfg.get('separator', '；')
            keywords_text = separator.join(keywords)
            content_run = kw_para.add_run()
            self.style_manager.set_mixed_font(
                content_run,
                keywords_text,
                chinese_font=kw_cfg.get('content_font'),
                english_font=self.style_manager.get_fonts().get('english'),
                size=kw_cfg.get('content_size'),
            )
            self.style_manager.apply_paragraph_style(kw_para, kw_cfg)

    def _generate_abstract_en(self, abstract_data):
        config = self.style_manager.config.get('abstract_en', {})
        title_cfg = config.get('title', {})
        title_text = title_cfg.get('text', 'ABSTRACT')
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_text)
        self.style_manager.set_mixed_font(
            title_run,
            title_text,
            chinese_font=title_cfg.get('font'),
            english_font=title_cfg.get('font'),
            size=title_cfg.get('size'),
            bold=title_cfg.get('bold', True),
        )
        self.style_manager.apply_paragraph_style(title_para, title_cfg)

        body_cfg = config.get('content', {})
        for para_text in abstract_data.get('content', []):
            para = self.doc.add_paragraph()
            run = para.add_run(para_text)
            run.font.name = body_cfg.get('font', 'Times New Roman')
            if body_cfg.get('size'):
                run.font.size = Pt(body_cfg['size'])
            self.style_manager.apply_paragraph_style(para, body_cfg)

        keywords = abstract_data.get('keywords') or []
        if keywords:
            kw_cfg = config.get('keywords', {})
            kw_para = self.doc.add_paragraph()
            label_text = kw_cfg.get('label_text', 'Keywords:')
            label_run = kw_para.add_run(f'{label_text} ')
            label_run.font.name = kw_cfg.get('label_font', 'Times New Roman')
            if kw_cfg.get('label_size'):
                label_run.font.size = Pt(kw_cfg['label_size'])
            label_run.font.bold = kw_cfg.get('label_bold', True)
            formatted_keywords = self._apply_keyword_rule(keywords, kw_cfg.get('casing_rule'))
            content_run = kw_para.add_run(kw_cfg.get('separator', '; ').join(formatted_keywords))
            content_run.font.name = kw_cfg.get('content_font', 'Times New Roman')
            if kw_cfg.get('content_size'):
                content_run.font.size = Pt(kw_cfg['content_size'])
            self.style_manager.apply_paragraph_style(kw_para, kw_cfg)

    def _generate_toc(self, chapters, special_sections=None):
        toc_cfg = self.style_manager.get_toc_config()
        title_cfg = toc_cfg.get('title', {})
        title_text = title_cfg.get('text', '目 录')
        title_para = self.doc.add_paragraph()
        title_run = title_para.add_run(title_text)
        self.style_manager.set_mixed_font(
            title_run,
            title_text,
            chinese_font=title_cfg.get('font'),
            english_font=title_cfg.get('font'),
            size=title_cfg.get('size'),
            bold=title_cfg.get('bold', True),
        )
        self.style_manager.apply_paragraph_style(title_para, title_cfg)

        level_styles = toc_cfg.get('levels', {})
        for chapter in chapters:
            self._add_toc_entry_for_level(
                chapter,
                level_styles.get('level1', {}),
                display_text=f"{self.style_manager.format_heading_label(1, chapter.get('number', 1))} {chapter.get('title', '')}",
                bookmark=f"_Chapter_{chapter.get('number', '')}",
            )
            for item in chapter.get('content', []):
                if item['type'] == 'heading2':
                    text = f"{self.style_manager.format_heading_label(2, chapter.get('number', 1), item.get('ordinal', 1))} {item['text']}"
                    bookmark = self._build_heading_bookmark(2, chapter.get('number', 1), item.get('ordinal', 1))
                    self._add_toc_entry_for_level(
                        chapter,
                        level_styles.get('level2', {}),
                        display_text=text,
                        bookmark=bookmark,
                    )
                elif item['type'] == 'heading3':
                    text = f"{self.style_manager.format_heading_label(3, chapter.get('number', 1), item.get('ordinal', 1))} {item['text']}"
                    bookmark = self._build_heading_bookmark(3, chapter.get('number', 1), item.get('ordinal', 1), item.get('parent'))
                    self._add_toc_entry_for_level(
                        chapter,
                        level_styles.get('level3', {}),
                        display_text=text,
                        bookmark=bookmark,
                    )

        for section in special_sections or []:
            self._add_toc_entry_for_level(
                None,
                level_styles.get('level1', {}),
                display_text=section.get('title', ''),
                bookmark=section.get('bookmark'),
            )

        self.audit_log['toc_inserted'] = True

    def _generate_chapter(self, chapter, chapter_idx):
        h1_style = self.style_manager.get_heading_style(1)
        para = self.doc.add_paragraph()
        chapter_num = chapter.get('number', chapter_idx + 1)
        label = self.style_manager.format_heading_label(1, chapter_num)
        label_run = para.add_run(f'{label} ')
        label_font = h1_style.get('number_font', h1_style.get('font'))
        label_run.font.name = label_font
        label_run._element.get_or_add_rPr()
        label_run.font.size = Pt(h1_style['size'])
        label_run.font.bold = h1_style.get('bold', True)

        title_run = para.add_run(chapter.get('title', ''))
        self.style_manager.set_mixed_font(
            title_run,
            chapter.get('title', ''),
            chinese_font=h1_style.get('font'),
            english_font=self.style_manager.get_fonts().get('english'),
            size=h1_style.get('size'),
            bold=h1_style.get('bold', True),
        )
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if 'space_before' in h1_style:
            para.paragraph_format.space_before = Pt(h1_style['space_before'])
        self._add_bookmark_to_paragraph(para, f'_Chapter_{chapter_num}')
        self.audit_log['heading_runs_have_east_asia'] = True

        for item in chapter.get('content', []):
            if item['type'] == 'paragraph':
                self._add_paragraph(item['text'])
            elif item['type'] == 'heading2':
                self._add_heading2(chapter_num, item)
            elif item['type'] == 'heading3':
                self._add_heading3(chapter_num, item)
            elif item['type'] == 'figure':
                self._add_figure(item)
            elif item['type'] == 'table':
                self._add_table(item)
            elif item['type'] == 'formula':
                self._add_formula(item)

    def _add_paragraph(self, text):
        if not text or not text.strip():
            return
        para_style = self.style_manager.get_paragraph_style()
        para = self.doc.add_paragraph()
        run = para.add_run()
        self.style_manager.set_mixed_font(
            run,
            text,
            chinese_font=para_style.get('font'),
            english_font=self.style_manager.get_fonts().get('english'),
            size=para_style.get('size'),
        )
        self.style_manager.apply_paragraph_style(para, para_style)
        self.audit_log['paragraph_spacing_applied'] = True

    def _add_heading2(self, chapter_num: int, item: Dict[str, Any]):
        style = self.style_manager.get_heading_style(2)
        para = self.doc.add_paragraph()
        label = self.style_manager.format_heading_label(2, chapter_num, item.get('ordinal', 1))
        label_run = para.add_run(f'{label} ')
        label_run.font.name = style.get('number_font', style.get('font'))
        label_run.font.size = Pt(style['size'])
        label_run.font.bold = style.get('bold', True)
        title_run = para.add_run(item['text'])
        self.style_manager.set_mixed_font(
            title_run,
            item['text'],
            chinese_font=style.get('font'),
            english_font=self.style_manager.get_fonts().get('english'),
            size=style.get('size'),
            bold=style.get('bold', True),
        )
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        bookmark = self._build_heading_bookmark(2, chapter_num, item.get('ordinal', 1))
        self._add_bookmark_to_paragraph(para, bookmark)
        self.audit_log['heading_runs_have_east_asia'] = True

    def _add_heading3(self, chapter_num: int, item: Dict[str, Any]):
        style = self.style_manager.get_heading_style(3)
        para = self.doc.add_paragraph()
        label = self.style_manager.format_heading_label(3, chapter_num, item.get('ordinal', 1))
        label_run = para.add_run(f'{label} ')
        label_run.font.name = style.get('number_font', style.get('font'))
        label_run.font.size = Pt(style['size'])
        label_run.font.bold = style.get('bold', True)
        title_run = para.add_run(item['text'])
        self.style_manager.set_mixed_font(
            title_run,
            item['text'],
            chinese_font=style.get('font'),
            english_font=self.style_manager.get_fonts().get('english'),
            size=style.get('size'),
            bold=style.get('bold', True),
        )
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        bookmark = self._build_heading_bookmark(3, chapter_num, item.get('ordinal', 1), item.get('parent'))
        self._add_bookmark_to_paragraph(para, bookmark)

    def _add_figure(self, figure_data: Dict[str, Any]):
        fig_cfg = self.style_manager.get_figure_style()
        chapter_num = self._chapter_from_number(figure_data.get('number', '1-1'))
        width_in = fig_cfg.get('width_in', 5)
        self.doc.add_paragraph()
        img_para = self.doc.add_paragraph()
        img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        image_path = figure_data.get('path')
        if image_path and os.path.exists(image_path):
            run = img_para.add_run()
            try:
                run.add_picture(image_path, width=Inches(width_in))
            except Exception as exc:
                placeholder = img_para.add_run(f'[图片加载失败: {exc}]')
                placeholder.font.color.rgb = RGBColor(200, 0, 0)
        else:
            placeholder = img_para.add_run(f"[图{figure_data.get('number', '')} 缺失]")
            placeholder.font.color.rgb = RGBColor(200, 0, 0)
            placeholder.font.size = Pt(fig_cfg['caption'].get('size', 12))
            self.audit_log['missing_figures'].append(figure_data.get('number', ''))

        caption_para = self.doc.add_paragraph()
        numbering_format = fig_cfg.get('numbering_format', '图{chapter}-{seq}')
        self._render_number_with_seq(caption_para, numbering_format, 'Figure', chapter_num)
        caption_para.add_run(fig_cfg['caption'].get('label_separator', ' '))
        caption_text = figure_data.get('caption', '')
        caption_para.add_run(caption_text)
        self._apply_caption_style(caption_para, fig_cfg.get('caption', {}))
        self.style_manager.apply_paragraph_style(caption_para, fig_cfg.get('caption', {}))

        if figure_data.get('source'):
            source_para = self.doc.add_paragraph()
            source_run = source_para.add_run(figure_data['source'])
            source_cfg = fig_cfg.get('source', {})
            self.style_manager.set_mixed_font(
                source_run,
                figure_data['source'],
                chinese_font=source_cfg.get('font'),
                english_font=self.style_manager.get_fonts().get('english'),
                size=source_cfg.get('size', 9),
            )

        self.doc.add_paragraph()

    def _add_table(self, table_data: Dict[str, Any]):
        tbl_cfg = self.style_manager.get_table_style()
        chapter_num = self._chapter_from_number(table_data.get('number', '1-1'))
        self.doc.add_paragraph()
        caption_para = self.doc.add_paragraph()
        numbering_format = tbl_cfg.get('number_format', '表{chapter}-{seq}')
        self._render_number_with_seq(caption_para, numbering_format, 'Table', chapter_num)
        caption_para.add_run(' ')
        caption_para.add_run(table_data.get('caption', ''))
        self._apply_caption_style(caption_para, tbl_cfg.get('caption', {}))
        self.style_manager.apply_paragraph_style(caption_para, tbl_cfg.get('caption', {}))
        caption_para.paragraph_format.keep_with_next = True

        rows = table_data.get('rows') or []
        if not rows:
            return
        table = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        for r_idx, row in enumerate(rows):
            for c_idx, cell_value in enumerate(row):
                cell = table.rows[r_idx].cells[c_idx]
                cell.text = cell_value
                for p in cell.paragraphs:
                    alignment = tbl_cfg.get('content_alignment', 'center')
                    p.alignment = {
                        'center': WD_ALIGN_PARAGRAPH.CENTER,
                        'left': WD_ALIGN_PARAGRAPH.LEFT,
                        'right': WD_ALIGN_PARAGRAPH.RIGHT,
                    }.get(alignment, WD_ALIGN_PARAGRAPH.CENTER)
                    for run in p.runs:
                        self.style_manager.set_mixed_font(
                            run,
                            run.text,
                            chinese_font=tbl_cfg.get('content_font'),
                            english_font=self.style_manager.get_fonts().get('english'),
                            size=tbl_cfg.get('content_size'),
                            bold=(r_idx == 0),
                        )
        self._repeat_table_header(table)
        self._set_table_borders(table, tbl_cfg)

        if table_data.get('source'):
            source_para = self.doc.add_paragraph()
            source_run = source_para.add_run(table_data['source'])
            source_cfg = tbl_cfg.get('source', {})
            self.style_manager.set_mixed_font(
                source_run,
                table_data['source'],
                chinese_font=source_cfg.get('font'),
                english_font=self.style_manager.get_fonts().get('english'),
                size=source_cfg.get('size', 9),
            )
        self.doc.add_paragraph()

    def _add_formula(self, formula_data: Dict[str, Any]):
        formula_style = self.style_manager.get_formula_style()
        formula_content = formula_data.get('content', '')
        formula_lines = [line.strip() for line in formula_content.split('\n') if line.strip()]
        if not formula_lines:
            return

        chapter_num = self._chapter_from_number(formula_data.get('number', '1'))
        numbering_template = formula_style.get('number_format', '({seq})')
        numbering_position = formula_style.get('numbering_position', 'right_same_line')
        seq_type = 'Equation'
        self.audit_log['omml_used'] = True

        spacer = self.doc.add_paragraph()
        spacer.paragraph_format.keep_with_next = True

        for line_idx, line in enumerate(formula_lines):
            paragraph = self.doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.keep_together = True
            is_last_line = line_idx == len(formula_lines) - 1

            if numbering_position == 'right_same_line' and is_last_line:
                self.style_manager.apply_tab_stop(paragraph, position_cm=8.25, alignment='center', leader=None)
                self.style_manager.apply_tab_stop(paragraph, position_cm=16.0, alignment='right', leader=None)
                paragraph.add_run('\t')

            try:
                oMathPara = OxmlElement('m:oMathPara')
                oMath = OxmlElement('m:oMath')
                self._build_omml_runs(oMath, line, formula_style)
                self._apply_math_justification(oMathPara, formula_style.get('alignment', 'center'))
                oMathPara.append(oMath)
                paragraph._element.append(oMathPara)
            except Exception:
                fallback_run = paragraph.add_run(line)
                fallback_run.font.name = formula_style.get('font', 'Times New Roman')
                fallback_run.font.size = Pt(formula_style.get('size', 12))

            if is_last_line:
                if numbering_position == 'right_same_line':
                    paragraph.add_run('\t')
                    runs = self._render_number_with_seq(paragraph, numbering_template, seq_type, chapter_num)
                    for run in runs:
                        run.font.name = formula_style.get('font', 'Times New Roman')
                        run.font.size = Pt(formula_style.get('size', 12))
                else:
                    num_para = self.doc.add_paragraph()
                    num_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    runs = self._render_number_with_seq(num_para, numbering_template, seq_type, chapter_num)
                    for run in runs:
                        run.font.name = formula_style.get('font', 'Times New Roman')
                        run.font.size = Pt(formula_style.get('size', 12))
        self.doc.add_paragraph()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _add_toc_entry_for_level(self, chapter, style_cfg: Dict[str, Any], display_text: str, bookmark: Optional[str]):
        if not bookmark or not display_text:
            return
        para = self.doc.add_paragraph()
        font_name = style_cfg.get('font', self.style_manager.get_fonts().get('chinese', '宋体'))
        font_size = style_cfg.get('size', 12)
        self.style_manager.apply_tab_stop(para, position_cm=16.0, alignment='right', leader='dot')
        if 'left_indent_chars' in style_cfg:
            indent = self._chars_to_pt(style_cfg['left_indent_chars'], font_size)
            if indent:
                para.paragraph_format.left_indent = indent
        self._insert_toc_hyperlink(
            para,
            display_text,
            bookmark,
            font_name=font_name,
            font_size=font_size,
            bold=style_cfg.get('bold', False),
        )
        para.add_run('\t')
        page_run = para.add_run()
        english_font = self.style_manager.get_fonts().get('english', 'Times New Roman')
        page_run.font.name = english_font
        page_run.font.size = Pt(font_size)
        self._add_pageref_field(page_run, bookmark)

    def _render_number_with_seq(self, paragraph, template: str, seq_type: str, chapter_num: int):
        runs: List[Any] = []
        tokens = re.split(r'(\{[^}]+\})', template)
        uses_chapter = '{chapter}' in template
        for token in tokens:
            if not token:
                continue
            if token == '{chapter}':
                paragraph.add_run(str(chapter_num))
                runs.append(paragraph.runs[-1])
            elif token == '{seq}':
                before = len(paragraph.runs)
                if uses_chapter:
                    self._add_chapter_based_seq_field(paragraph, seq_type, chapter_num)
                else:
                    self._add_seq_field(paragraph, seq_type)
                self.audit_log['seq_fields_used'] = True
                if len(paragraph.runs) > before:
                    runs.append(paragraph.runs[-1])
            else:
                paragraph.add_run(token)
                runs.append(paragraph.runs[-1])
        return runs

    def _apply_caption_style(self, paragraph, caption_cfg: Dict[str, Any]):
        font_name = caption_cfg.get('font', self.style_manager.get_fonts().get('chinese', '宋体'))
        font_size = caption_cfg.get('size', 12)
        for run in paragraph.runs:
            run.font.name = font_name
            run._element.get_or_add_rPr()
            rPr = run._element.rPr
            if rPr is not None:
                fonts = rPr.rFonts
                if fonts is None:
                    fonts = OxmlElement('w:rFonts')
                    rPr.append(fonts)
                english_font = self.style_manager.get_fonts().get('english', 'Times New Roman')
                fonts.set(qn('w:eastAsia'), font_name)
                fonts.set(qn('w:ascii'), english_font)
                fonts.set(qn('w:hAnsi'), english_font)
            run.font.size = Pt(font_size)
        paragraph.alignment = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
        }.get(caption_cfg.get('alignment', 'center'), WD_ALIGN_PARAGRAPH.CENTER)

    def _chapter_from_number(self, number_text: str) -> int:
        parts = re.split(r'[-.]', number_text)
        try:
            return int(parts[0])
        except (TypeError, ValueError):
            return 1

    def _build_heading_bookmark(self, level: int, chapter_num: int, ordinal: int, parent: Optional[int] = None) -> str:
        if level == 2:
            return f'_Heading_{chapter_num}_{ordinal}'
        if level == 3:
            return f'_Heading_{chapter_num}_{parent or 0}_{ordinal}'
        return f'_Heading_{chapter_num}'

    def _chars_to_pt(self, char_count: float, font_size: float) -> Optional[Pt]:
        if char_count is None or font_size is None:
            return None
        return Pt(float(char_count) * float(font_size))

    def _apply_keyword_rule(self, keywords: List[str], rule: Optional[str]) -> List[str]:
        if not rule:
            return keywords
        if rule == 'capitalize_each_word':
            return [' '.join(word.capitalize() for word in kw.split()) for kw in keywords]
        if rule == 'upper':
            return [kw.upper() for kw in keywords]
        if rule == 'lower':
            return [kw.lower() for kw in keywords]
        return keywords

    def _insert_toc_hyperlink(self, paragraph, text: str, bookmark: str, font_name: str, font_size: float, bold: bool):
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('w:anchor'), bookmark)
        hyperlink.set(qn('w:history'), '1')
        run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        fonts = OxmlElement('w:rFonts')
        english_font = self.style_manager.get_fonts().get('english', 'Times New Roman')
        fonts.set(qn('w:ascii'), english_font)
        fonts.set(qn('w:hAnsi'), english_font)
        fonts.set(qn('w:eastAsia'), font_name)
        rPr.append(fonts)
        size_element = OxmlElement('w:sz')
        size_element.set(qn('w:val'), str(int(float(font_size) * 2)))
        rPr.append(size_element)
        if bold:
            bold_element = OxmlElement('w:b')
            bold_element.set(qn('w:val'), '1')
            rPr.append(bold_element)
        color = OxmlElement('w:color')
        color.set(qn('w:val'), self.style_manager.get_link_color())
        rPr.append(color)
        run.append(rPr)
        text_element = OxmlElement('w:t')
        text_element.text = text
        run.append(text_element)
        hyperlink.append(run)
        paragraph._element.append(hyperlink)

    def _enable_field_updates(self):
        settings = self.doc.settings.element
        update_fields = settings.find(qn('w:updateFields'))
        if update_fields is None:
            update_fields = OxmlElement('w:updateFields')
            settings.append(update_fields)
        update_fields.set(qn('w:val'), 'true')
