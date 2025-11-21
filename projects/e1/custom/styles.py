from typing import Any, Dict, Optional

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from src.styles import StyleManager as BaseStyleManager


class E1StyleManager(BaseStyleManager):
    """
    Project-specific style helper that keeps every layout knob config-driven.
    """

    def __init__(self, config_path: str):
        super().__init__(config_path)
        self._fonts = self.config.get('fonts', {})
        self._size_map = self.config.get('sizes', {})
        self._default_para_style = self.config.get('body', {}).get('paragraph', {})
        self._link_color = '000000'

    # ------------------------------------------------------------------
    # Generic helpers
    # ------------------------------------------------------------------
    def get_fonts(self) -> Dict[str, str]:
        return self._fonts

    def get_link_color(self) -> str:
        return self._link_color

    def get_toc_config(self) -> Dict[str, Any]:
        return self.config.get('toc', {})

    def get_document_settings(self) -> Dict[str, Any]:
        document_config = self.config.get('document', {}).copy()
        margins = document_config.get('margins', {})
        document_config['margins'] = {
            'top': margins.get('top', 2.5),
            'bottom': margins.get('bottom', 2.5),
            'left': margins.get('left', 3.0),
            'right': margins.get('right', 2.5),
        }
        return document_config

    # ------------------------------------------------------------------
    # Formatting utilities
    # ------------------------------------------------------------------
    def apply_paragraph_style(self, paragraph, style_config: Optional[Dict[str, Any]]):
        if not style_config:
            return

        fmt = paragraph.paragraph_format
        alignment = style_config.get('alignment')
        if alignment:
            paragraph.alignment = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
            }.get(alignment, WD_ALIGN_PARAGRAPH.LEFT)

        space_before = self._to_pt(style_config.get('space_before') or style_config.get('spacing_before_pt'))
        if space_before is not None:
            fmt.space_before = space_before

        space_after = self._to_pt(style_config.get('space_after') or style_config.get('spacing_after_pt'))
        if space_after is not None:
            fmt.space_after = space_after

        self._apply_line_spacing(fmt, style_config.get('line_spacing'))

        size_hint = self._resolve_size_value(style_config.get('size'))
        if size_hint is None:
            size_hint = self._resolve_size_value(self._default_para_style.get('size'))

        if 'first_line_indent' in style_config:
            indent = self._chars_to_pt(style_config['first_line_indent'], size_hint)
            if indent is not None:
                fmt.first_line_indent = indent

        if 'first_line_indent_chars' in style_config:
            indent = self._chars_to_pt(style_config['first_line_indent_chars'], size_hint)
            if indent is not None:
                fmt.first_line_indent = indent

        if 'left_indent_chars' in style_config:
            indent = self._chars_to_pt(style_config['left_indent_chars'], size_hint)
            if indent is not None:
                fmt.left_indent = indent

        if 'hanging_indent_chars' in style_config:
            indent = self._chars_to_pt(style_config['hanging_indent_chars'], size_hint)
            if indent is not None:
                fmt.left_indent = indent
                fmt.first_line_indent = Pt(-indent.pt)
                self._write_char_based_indent(paragraph, style_config['hanging_indent_chars'])

        if 'first_line_chars' in style_config:
            indent = self._chars_to_pt(style_config['first_line_chars'], size_hint)
            if indent is not None:
                fmt.first_line_indent = indent

    def set_mixed_font(
        self,
        run,
        text: str,
        chinese_font: Optional[str] = None,
        english_font: Optional[str] = None,
        size: Optional[float] = None,
        bold: bool = False,
        italic: bool = False,
    ):
        if text is None:
            return
        run.text = text
        ascii_font = english_font or self._fonts.get('english', 'Times New Roman')
        east_font = chinese_font or self._fonts.get('chinese', '宋体')
        run.font.name = ascii_font
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.rFonts
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)
        r_fonts.set(qn('w:ascii'), ascii_font)
        r_fonts.set(qn('w:hAnsi'), ascii_font)
        r_fonts.set(qn('w:eastAsia'), east_font)
        if size:
            run.font.size = self._to_pt(size)
        if bold:
            run.font.bold = True
        if italic:
            run.font.italic = True

    def apply_tab_stop(self, paragraph, position_cm=16.0, alignment='right', leader='dot'):
        pPr = paragraph._element.get_or_add_pPr()
        tabs = pPr.find(qn('w:tabs'))
        if tabs is None:
            tabs = OxmlElement('w:tabs')
            pPr.append(tabs)
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), alignment)
        tab.set(qn('w:pos'), str(int(position_cm * 567)))
        if leader:
            tab.set(qn('w:leader'), leader)
        tabs.append(tab)

    # ------------------------------------------------------------------
    # Heading number helpers
    # ------------------------------------------------------------------
    def format_heading_label(self, level: int, chapter_number: int, ordinal: int = 0) -> str:
        if level == 1:
            return f'第{self._to_chinese(chapter_number)}章'
        if level == 2:
            return f'第{self._to_chinese(max(ordinal, 1))}节'
        if level == 3:
            return f'{self._to_chinese(max(ordinal, 1))}、'
        return ''

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _resolve_size_value(self, value: Any) -> Optional[float]:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            if value.endswith('pt'):
                try:
                    return float(value[:-2])
                except ValueError:
                    return None
            if value in self._size_map:
                return float(self._size_map[value])
            try:
                return float(value)
            except ValueError:
                return None
        return None

    def _to_pt(self, value: Any) -> Optional[Pt]:
        numeric = self._resolve_size_value(value)
        if numeric is None:
            return None
        return Pt(numeric)

    def _chars_to_pt(self, char_count: float, font_size_pt: Optional[float]) -> Optional[Pt]:
        if char_count is None or font_size_pt is None:
            return None
        return Pt(float(char_count) * float(font_size_pt))

    def _apply_line_spacing(self, paragraph_format, spacing_value: Any):
        if spacing_value is None:
            return
        if isinstance(spacing_value, str):
            if spacing_value.startswith('fixed_'):
                numeric = spacing_value.split('_', 1)[1].replace('pt', '')
                try:
                    paragraph_format.line_spacing = Pt(float(numeric))
                except ValueError:
                    pass
                return
            if spacing_value.startswith('multiple_'):
                try:
                    multiple = float(spacing_value.split('_', 1)[1])
                    paragraph_format.line_spacing = multiple
                    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                except ValueError:
                    pass
                return
        try:
            numeric = float(spacing_value)
        except (TypeError, ValueError):
            return
        rule_map = {
            1.0: WD_LINE_SPACING.SINGLE,
            1.5: WD_LINE_SPACING.ONE_POINT_FIVE,
            2.0: WD_LINE_SPACING.DOUBLE,
        }
        if numeric in rule_map:
            paragraph_format.line_spacing_rule = rule_map[numeric]
        elif numeric <= 3:
            paragraph_format.line_spacing = numeric
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        else:
            paragraph_format.line_spacing = Pt(numeric)

    def _write_char_based_indent(self, paragraph, char_count: float):
        pPr = paragraph._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        ind.set(qn('w:firstLineChars'), str(int(char_count * 100)))

    def _to_chinese(self, number: int) -> str:
        digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
        units = ['', '十', '百', '千']
        if number <= 0:
            return digits[0]
        if number < 10:
            return digits[number]
        if number < 20:
            tail = number % 10
            prefix = '十'
            return prefix if tail == 0 else f'{prefix}{digits[tail]}'
        chars = []
        str_num = str(number)
        length = len(str_num)
        for idx, ch in enumerate(str_num):
            digit = int(ch)
            pos = length - idx - 1
            if digit == 0:
                if chars and chars[-1] != digits[0]:
                    chars.append(digits[0])
                continue
            chars.append(digits[digit])
            chars.append(units[pos])
        result = ''.join(chars).rstrip('零')
        result = result.replace('零零', '零')
        if result.endswith('零'):
            result = result[:-1]
        return result or digits[0]
