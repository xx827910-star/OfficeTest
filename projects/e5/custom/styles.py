"""Config-driven style manager for project e5."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Optional

from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


class StyleManager:
    """Load thesis_format.json and expose helpers for formatting."""

    def __init__(self, config_path: Path):
        config_path = Path(config_path)
        self.config = json.loads(config_path.read_text(encoding='utf-8'))

    # ------------------------------------------------------------------
    # Document-level helpers
    # ------------------------------------------------------------------
    def configure_document_styles(self, document) -> None:
        """Apply Normal/Heading/TOC/Hyperlink styles based on config."""
        fonts = self.get_fonts()
        paragraph_cfg = self.get_paragraph_style()

        normal = document.styles['Normal']
        normal.font.name = fonts.get('english', 'Times New Roman')
        normal.font.size = Pt(paragraph_cfg.get('size', 12))
        self._set_style_fonts(normal, fonts)
        self._apply_paragraph_spacing(normal.paragraph_format, paragraph_cfg)

        for level in (1, 2, 3):
            self._configure_heading_style(document, level)

        self._configure_toc_styles(document)
        self._configure_hyperlink_style(document)

    def _configure_heading_style(self, document, level: int) -> None:
        heading_style = document.styles[f'Heading {level}']
        cfg = self.get_heading_style(level)
        heading_style.font.name = cfg.get('font', '宋体')
        heading_style.font.size = Pt(cfg.get('size', 14))
        heading_style.font.bold = cfg.get('bold', False)
        self._set_style_fonts(
            heading_style,
            {
                'english': cfg.get('number_font', cfg.get('font', 'Times New Roman')),
                'chinese': cfg.get('font', '宋体')
            }
        )
        alignment = cfg.get('alignment', 'left').lower()
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        heading_style.paragraph_format.alignment = alignment_map.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self._apply_paragraph_spacing(heading_style.paragraph_format, cfg)
        self._apply_line_spacing(heading_style.paragraph_format, cfg)

    def _configure_toc_styles(self, document) -> None:
        toc_cfg = self.get_toc_style()
        entry_cfg = toc_cfg.get('entry', {})
        levels_cfg = toc_cfg.get('levels', {})
        fonts = self.get_fonts()
        for level in (1, 2, 3):
            style_name = f'TOC {level}'
            if style_name not in document.styles:
                continue
            style = document.styles[style_name]
            cn_font = entry_cfg.get('font_chinese', fonts.get('chinese', '宋体'))
            en_font = entry_cfg.get('font_english', fonts.get('english', 'Times New Roman'))
            style.font.name = en_font
            style.font.size = Pt(entry_cfg.get('size', 12))
            style.font.bold = levels_cfg.get(f'heading{level}', {}).get('bold', False)
            self._set_style_fonts(style, {'english': en_font, 'chinese': cn_font})
            pf = style.paragraph_format
            pf.space_before = Pt(entry_cfg.get('space_before', 0))
            pf.space_after = Pt(entry_cfg.get('space_after', 0))
            indent_chars = levels_cfg.get(f'heading{level}', {}).get('indent_chars', 0)
            pf.left_indent = Pt(indent_chars * entry_cfg.get('size', 12))
            pf.tab_stops.clear_all()
            pf.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

    def _configure_hyperlink_style(self, document) -> None:
        try:
            hyperlink = document.styles['Hyperlink']
        except KeyError:
            return
        hyperlink.font.color.rgb = RGBColor(0, 0, 0)
        hyperlink.font.underline = False
        hyperlink.font.bold = False

    def _set_style_fonts(self, style, fonts: Dict[str, str]) -> None:
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        if 'chinese' in fonts:
            rFonts.set(qn('w:eastAsia'), fonts['chinese'])
        if 'english' in fonts:
            rFonts.set(qn('w:ascii'), fonts['english'])
            rFonts.set(qn('w:hAnsi'), fonts['english'])

    # ------------------------------------------------------------------
    # Apply styles to paragraph/run
    # ------------------------------------------------------------------
    def apply_paragraph_style(self, paragraph, style_config: Dict) -> None:
        if not style_config:
            return
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if style_config.get('alignment'):
            paragraph.alignment = alignment_map.get(style_config['alignment'], WD_ALIGN_PARAGRAPH.LEFT)
        self._apply_paragraph_spacing(paragraph.paragraph_format, style_config)
        self._apply_line_spacing(paragraph.paragraph_format, style_config)

        font_size = style_config.get('size', self.get_paragraph_style().get('size', 12))
        if 'hanging_indent_chars' in style_config:
            chars = style_config['hanging_indent_chars']
            paragraph.paragraph_format.left_indent = Pt(chars * font_size)
            paragraph.paragraph_format.first_line_indent = Pt(-chars * font_size)
            self._apply_character_indents(paragraph, first_line_chars=-chars, left_chars=chars)
        elif 'first_line_indent' in style_config:
            chars = style_config['first_line_indent']
            paragraph.paragraph_format.first_line_indent = Pt(chars * font_size)
            self._apply_character_indents(paragraph, first_line_chars=chars)

    def apply_run_style(self, run, style_config: Optional[Dict]) -> None:
        if not style_config:
            return
        cn_font = style_config.get('font_chinese') or style_config.get('font') or self.get_fonts().get('chinese', '宋体')
        en_font = style_config.get('font_english') or style_config.get('font') or self.get_fonts().get('english', 'Times New Roman')
        size = style_config.get('size') or self.get_paragraph_style().get('size', 12)
        self.set_mixed_font(
            run,
            run.text,
            chinese_font=cn_font,
            english_font=en_font,
            size=size,
            bold=style_config.get('bold', False),
            italic=style_config.get('italic', False),
            color=style_config.get('color')
        )

    def set_mixed_font(
        self,
        run,
        text: Optional[str],
        chinese_font: str,
        english_font: str,
        size: float,
        bold: bool = False,
        italic: bool = False,
        color: Optional[str] = None
    ) -> None:
        if text is not None:
            run.text = text
        run.font.name = english_font
        run.font.size = Pt(size)
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), english_font)
        rFonts.set(qn('w:hAnsi'), english_font)
        rFonts.set(qn('w:eastAsia'), chinese_font)

        if bold:
            run.font.bold = True
            self._ensure_bool_tag(rPr, 'w:b', True)
        else:
            self._ensure_bool_tag(rPr, 'w:b', False)

        if italic:
            run.font.italic = True
            self._ensure_bool_tag(rPr, 'w:i', True)
        else:
            self._ensure_bool_tag(rPr, 'w:i', False)

        if color:
            rgb = RGBColor.from_string(color.replace('#', '')) if isinstance(color, str) else RGBColor(0, 0, 0)
            run.font.color.rgb = rgb

    def _ensure_bool_tag(self, rPr, tag: str, enabled: bool) -> None:
        element = rPr.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            rPr.append(element)
        element.set(qn('w:val'), '1' if enabled else '0')

    # ------------------------------------------------------------------
    # Utility helpers
    # ------------------------------------------------------------------
    def _apply_paragraph_spacing(self, paragraph_format, cfg: Dict) -> None:
        if cfg.get('space_before') is not None:
            paragraph_format.space_before = Pt(cfg.get('space_before', 0))
        if cfg.get('space_after') is not None:
            paragraph_format.space_after = Pt(cfg.get('space_after', 0))

    def _apply_line_spacing(self, paragraph_format, cfg: Dict) -> None:
        if cfg.get('line_spacing_rule') == 'fixed':
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(cfg.get('line_spacing_pt', 20))
        elif cfg.get('line_spacing_rule') == 'multiple':
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            paragraph_format.line_spacing = cfg.get('line_spacing', 1.0)
        elif cfg.get('line_spacing'):
            rule_map = {
                1.0: WD_LINE_SPACING.SINGLE,
                1.5: WD_LINE_SPACING.ONE_POINT_FIVE,
                2.0: WD_LINE_SPACING.DOUBLE
            }
            paragraph_format.line_spacing_rule = rule_map.get(cfg['line_spacing'], WD_LINE_SPACING.SINGLE)
        elif cfg.get('line_spacing_pt'):
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            paragraph_format.line_spacing = Pt(cfg['line_spacing_pt'])

    def _apply_character_indents(self, paragraph, first_line_chars: float = 0.0, left_chars: float = 0.0) -> None:
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        if first_line_chars:
            ind.set(qn('w:firstLineChars'), str(int(first_line_chars * 100)))
        if left_chars:
            ind.set(qn('w:leftChars'), str(int(left_chars * 100)))

    # ------------------------------------------------------------------
    # Config accessors
    # ------------------------------------------------------------------
    def get_fonts(self) -> Dict:
        return self.config.get('fonts', {})

    def get_document_settings(self) -> Dict:
        return self.config.get('document', {})

    def get_page_number_config(self, section_name: str) -> Dict:
        mapping = {
            'abstract': self.config.get('abstract', {}),
            'toc': self.config.get('toc', {}),
            'body': self.config.get('body', {}),
            'references': self.config.get('references', {}),
            'acknowledgements': self.config.get('acknowledgements', {}),
            'appendix': self.config.get('appendix', {})
        }
        section = mapping.get(section_name, {})
        return section.get('page_number', {})

    def get_abstract_title_style(self) -> Dict:
        return self.config.get('abstract', {}).get('title', {})

    def get_abstract_content_style(self) -> Dict:
        return self.config.get('abstract', {}).get('content', {})

    def get_abstract_keywords_style(self) -> Dict:
        return self.config.get('abstract', {}).get('keywords', {})

    def get_toc_style(self) -> Dict:
        return self.config.get('toc', {})

    def get_heading_style(self, level: int) -> Dict:
        return self.config.get('body', {}).get({1: 'heading1', 2: 'heading2', 3: 'heading3'}[level], {})

    def get_paragraph_style(self) -> Dict:
        return self.config.get('body', {}).get('paragraph', {})

    def get_figure_style(self) -> Dict:
        return self.config.get('figure', {})

    def get_table_style(self) -> Dict:
        return self.config.get('table', {})

    def get_formula_style(self) -> Dict:
        return self.config.get('formula', {})

    def get_references_style(self) -> Dict:
        return self.config.get('references', {})

    def get_acknowledgement_style(self) -> Dict:
        return self.config.get('acknowledgements', {})

    def get_appendix_style(self) -> Dict:
        return self.config.get('appendix', {})

    def get_sizes(self) -> Dict:
        return self.config.get('sizes', {})
