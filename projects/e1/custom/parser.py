import os
import re
from glob import glob
from typing import Any, Dict, List, Optional


class E1ContentParser:
    """
    Lightweight structured parser for normalized thesis text tailored to project e1.
    """

    def __init__(self, image_dirs: Optional[List[str]] = None):
        default_dirs = [
            'projects/e1/input/images',
            'projects/e1/input/figures',
            'examples/extracted/e1/images',
            'examples/images',
        ]
        selected = image_dirs or default_dirs
        self.image_dirs = [os.path.abspath(p) for p in selected if p and os.path.exists(p)]
        self.missing_assets: List[Dict[str, str]] = []

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def parse_file(self, file_path: str) -> Dict[str, Any]:
        with open(file_path, 'r', encoding='utf-8') as handle:
            text = handle.read()
        return self.parse_text(text)

    def parse_text(self, text: str) -> Dict[str, Any]:
        lines = text.splitlines()
        content = {
            'title': '',
            'subtitle': '',
            'abstract': {'content': [], 'keywords': []},
            'abstract_en': {'content': [], 'keywords': []},
            'chapters': [],
            'references': [],
            'acknowledgements': [],
            'appendix': [],
        }

        current_section = None
        current_chapter = None
        current_paragraph: List[str] = []
        references_buffer: List[str] = []
        current_reference_lines: List[str] = []
        heading2_index = 0
        heading3_index = 0
        title_captured = False

        def flush_paragraph():
            nonlocal current_paragraph
            if not current_paragraph:
                return
            paragraph_text = ' '.join(segment.strip() for segment in current_paragraph if segment.strip()).strip()
            current_paragraph = []
            if not paragraph_text:
                return
            if current_section == 'abstract':
                content['abstract']['content'].append(paragraph_text)
            elif current_section == 'abstract_en':
                content['abstract_en']['content'].append(paragraph_text)
            elif current_section == 'body' and current_chapter:
                current_chapter['content'].append({'type': 'paragraph', 'text': paragraph_text})
            elif current_section == 'acknowledgements':
                content['acknowledgements'].append(paragraph_text)
            elif current_section == 'appendix':
                content['appendix'].append(paragraph_text)

        i = 0
        total_lines = len(lines)
        while i < total_lines:
            raw_line = lines[i]
            line = raw_line.strip()

            if not title_captured and line:
                content['title'] = line
                title_captured = True
                if i + 1 < total_lines and lines[i + 1].strip().startswith('——'):
                    content['subtitle'] = lines[i + 1].strip().lstrip('—').strip()
                    i += 2
                    continue
                i += 1
                continue

            if not line:
                flush_paragraph()
                i += 1
                continue

            if line == '[TOC]':
                i += 1
                while i < total_lines and lines[i].strip() != '[/TOC]':
                    i += 1
                i += 1
                continue

            if line == '[ABSTRACT]':
                current_section = 'abstract'
                current_paragraph = []
                i += 1
                continue

            if line == '[/ABSTRACT]':
                flush_paragraph()
                current_section = None
                i += 1
                continue

            if line == '[ABSTRACT_EN]':
                current_section = 'abstract_en'
                current_paragraph = []
                i += 1
                continue

            if line == '[/ABSTRACT_EN]':
                flush_paragraph()
                current_section = None
                i += 1
                continue

            if self._is_keywords_line(line):
                flush_paragraph()
                keywords = self._normalize_keywords(line)
                if current_section == 'abstract':
                    content['abstract']['keywords'] = keywords
                elif current_section == 'abstract_en':
                    content['abstract_en']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            if line == '[BODY]':
                current_section = 'body'
                i += 1
                continue

            if line == '[/BODY]':
                flush_paragraph()
                current_section = None
                i += 1
                continue

            if self._is_references_header(line):
                flush_paragraph()
                current_section = 'references'
                current_reference_lines = []
                i += 1
                continue

            if current_section == 'references':
                if line.startswith('[/REFERENCES]'):
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                    current_section = None
                    i += 1
                    continue
                if self._is_new_reference_entry(line, bool(current_reference_lines)):
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = [line]
                else:
                    current_reference_lines.append(line)
                i += 1
                continue

            if self._is_acknowledgements_header(line):
                flush_paragraph()
                current_section = 'acknowledgements'
                i += 1
                continue

            if line in ('[/ACKNOWLEDGEMENTS]', '[/致谢]'):
                flush_paragraph()
                current_section = None
                i += 1
                continue

            if self._is_appendix_header(line):
                flush_paragraph()
                current_section = 'appendix'
                i += 1
                continue

            if line in ('[/APPENDIX]', '[/附录]'):
                flush_paragraph()
                current_section = None
                i += 1
                continue

            chapter_match = self._match_level1_heading(line)
            if chapter_match:
                flush_paragraph()
                current_section = 'body'
                heading2_index = 0
                heading3_index = 0
                chapter_num, chapter_title = chapter_match
                current_chapter = {
                    'number': chapter_num,
                    'title': chapter_title,
                    'content': [],
                }
                content['chapters'].append(current_chapter)
                i += 1
                continue

            if current_section == 'body' and current_chapter:
                heading2_title = self._match_level2_heading(line)
                if heading2_title:
                    flush_paragraph()
                    heading2_index += 1
                    heading3_index = 0
                    current_chapter['content'].append({
                        'type': 'heading2',
                        'text': heading2_title,
                        'ordinal': heading2_index,
                    })
                    i += 1
                    continue

                heading3_title = self._match_level3_heading(line)
                if heading3_title:
                    flush_paragraph()
                    heading3_index += 1
                    current_chapter['content'].append({
                        'type': 'heading3',
                        'text': heading3_title,
                        'ordinal': heading3_index,
                        'parent': heading2_index,
                    })
                    i += 1
                    continue

                if line.startswith('[FIGURE:'):
                    flush_paragraph()
                    figure_block = self._parse_figure_block(lines, i)
                    i = figure_block['next_index']
                    figure_entry = figure_block['payload']
                    if figure_entry and current_chapter:
                        current_chapter['content'].append(figure_entry)
                    continue

                if line.startswith('[TABLE:'):
                    flush_paragraph()
                    table_block = self._parse_table_block(lines, i)
                    i = table_block['next_index']
                    if table_block['payload'] and current_chapter:
                        current_chapter['content'].append(table_block['payload'])
                    continue

                if line.startswith('[FORMULA:'):
                    flush_paragraph()
                    formula_block = self._parse_formula_block(lines, i)
                    i = formula_block['next_index']
                    if formula_block['payload'] and current_chapter:
                        current_chapter['content'].append(formula_block['payload'])
                    continue

            if current_section in {'abstract', 'abstract_en', 'body', 'acknowledgements', 'appendix'}:
                current_paragraph.append(line)

            i += 1

        flush_paragraph()
        if current_reference_lines:
            references_buffer.append(' '.join(current_reference_lines).strip())
        content['references'] = self._build_references_list(references_buffer)
        return content

    def get_missing_assets(self) -> List[Dict[str, str]]:
        return self.missing_assets

    # ------------------------------------------------------------------
    # Block parsers
    # ------------------------------------------------------------------
    def _parse_figure_block(self, lines: List[str], start_index: int) -> Dict[str, Any]:
        match = re.match(r'\[FIGURE:([\d\-]+)\]', lines[start_index].strip())
        figure_number = match.group(1) if match else '1-1'
        i = start_index + 1
        caption = ''
        source = ''
        hint = None
        if i < len(lines):
            parts = [part.strip() for part in lines[i].split('|')]
            caption = parts[0] if parts else ''
            if len(parts) > 1:
                source = parts[1]
            if len(parts) > 2:
                hint = parts[2]
            i += 1
        while i < len(lines) and lines[i].strip() != '[/FIGURE]':
            i += 1
        next_index = min(i + 1, len(lines))
        image_path = self._resolve_asset_path(figure_number, hint)
        if not image_path:
            self.missing_assets.append({'type': 'figure', 'id': figure_number, 'hint': hint or ''})
        payload = {
            'type': 'figure',
            'number': figure_number,
            'caption': caption,
            'source': source,
            'path': image_path,
        }
        return {'next_index': next_index, 'payload': payload}

    def _parse_table_block(self, lines: List[str], start_index: int) -> Dict[str, Any]:
        match = re.match(r'\[TABLE:([\d\-]+)\]', lines[start_index].strip())
        table_number = match.group(1) if match else '1-1'
        i = start_index + 1
        caption = ''
        source = ''
        rows: List[List[str]] = []
        if i < len(lines):
            parts = [part.strip() for part in lines[i].split('|')]
            caption = parts[0] if parts else ''
            if len(parts) > 1:
                source = parts[1]
            i += 1
        while i < len(lines):
            line = lines[i].strip()
            if line == '[/TABLE]':
                break
            if line:
                rows.append([cell.strip() for cell in line.split('|')])
            i += 1
        next_index = min(i + 1, len(lines))
        payload = {
            'type': 'table',
            'number': table_number,
            'caption': caption,
            'source': source,
            'rows': rows,
        }
        return {'next_index': next_index, 'payload': payload}

    def _parse_formula_block(self, lines: List[str], start_index: int) -> Dict[str, Any]:
        match = re.match(r'\[FORMULA:([\d\-]+)\]', lines[start_index].strip())
        formula_number = match.group(1) if match else '1'
        i = start_index + 1
        formula_lines: List[str] = []
        while i < len(lines):
            line = lines[i].strip()
            if line == '[/FORMULA]':
                break
            if line:
                formula_lines.append(line)
            i += 1
        next_index = min(i + 1, len(lines))
        payload = {
            'type': 'formula',
            'number': formula_number,
            'content': '\n'.join(formula_lines),
        }
        return {'next_index': next_index, 'payload': payload}

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _match_level1_heading(self, line: str) -> Optional[tuple]:
        pattern_digit = re.match(r'^第([一二三四五六七八九十百千\d]+)章[：:\s]*(.+)$', line)
        if pattern_digit:
            return self._chinese_to_arabic(pattern_digit.group(1)), pattern_digit.group(2).strip()
        pattern_chinese = re.match(r'^([一二三四五六七八九十百千]+)[、.．]\s*(.+)$', line)
        if pattern_chinese:
            return self._chinese_to_arabic(pattern_chinese.group(1)), pattern_chinese.group(2).strip()
        return None

    def _match_level2_heading(self, line: str) -> Optional[str]:
        pattern = re.match(r'^（([一二三四五六七八九十百千]+)）\s*(.+)$', line)
        if pattern:
            return pattern.group(2).strip()
        pattern_digit = re.match(r'^第([一二三四五六七八九十百千\d]+)节[：:\s]*(.+)$', line)
        if pattern_digit:
            return pattern_digit.group(2).strip()
        return None

    def _match_level3_heading(self, line: str) -> Optional[str]:
        pattern_digit = re.match(r'^(\d+)[\.\s]+(.+)$', line)
        if pattern_digit:
            return pattern_digit.group(2).strip()
        pattern_parenthesis = re.match(r'^（\d+）\s*(.+)$', line)
        if pattern_parenthesis:
            return pattern_parenthesis.group(1).strip()
        return None

    def _is_keywords_line(self, line: str) -> bool:
        lowered = line.lower()
        return lowered.startswith('关键词') or lowered.startswith('key words') or lowered.startswith('keywords')

    def _normalize_keywords(self, line: str) -> List[str]:
        parts = re.split('[:：]', line, 1)
        payload = parts[1] if len(parts) > 1 else parts[0]
        keywords = [token.strip() for token in re.split('[;；,，]', payload) if token.strip()]
        return keywords

    def _resolve_asset_path(self, figure_number: str, filename_hint: Optional[str]) -> Optional[str]:
        candidates: List[str] = []
        if filename_hint:
            hint_path = filename_hint
            if not os.path.isabs(hint_path):
                for base_dir in self.image_dirs:
                    candidates.append(os.path.join(base_dir, hint_path))
            else:
                candidates.append(hint_path)
        normalized = figure_number.replace('-', '_')
        extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.webp']
        for base_dir in self.image_dirs:
            for ext in extensions:
                candidates.append(os.path.join(base_dir, f'figure_{normalized}{ext}'))
                candidates.append(os.path.join(base_dir, f'image_{normalized}{ext}'))
        seen = set()
        for candidate in candidates:
            if not candidate or candidate in seen:
                continue
            seen.add(candidate)
            if os.path.exists(candidate):
                return candidate
        for base_dir in self.image_dirs:
            pattern = os.path.join(base_dir, f'figure_{normalized}.*')
            matches = sorted(glob(pattern))
            if matches:
                return matches[0]
        return None

    def _chinese_to_arabic(self, chinese_num: str) -> int:
        digits = {'零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
        units = {'十': 10, '百': 100, '千': 1000}
        if chinese_num.isdigit():
            return int(chinese_num)
        total = 0
        unit = 1
        for char in reversed(chinese_num):
            if char in units:
                unit = units[char]
                if total == 0:
                    total = unit
            elif char in digits:
                total += digits[char] * unit
            else:
                continue
        return total or 1

    def _is_references_header(self, line: str) -> bool:
        normalized = line.strip('[]（）()【】')
        return normalized.upper() == 'REFERENCES' or normalized == '参考文献'

    def _is_new_reference_entry(self, line: str, has_current: bool) -> bool:
        if not has_current:
            return True
        return bool(re.match(r'^\[\d+\]', line))

    def _is_acknowledgements_header(self, line: str) -> bool:
        normalized = line.strip('[]（）()【】')
        return normalized.upper() == 'ACKNOWLEDGEMENTS' or normalized == '致谢'

    def _is_appendix_header(self, line: str) -> bool:
        normalized = line.strip('[]（）()【】')
        return normalized.upper() == 'APPENDIX' or normalized == '附录'

    def _build_references_list(self, raw_entries: List[str]) -> List[Dict[str, str]]:
        references: List[Dict[str, str]] = []
        for idx, entry in enumerate(raw_entries, 1):
            text = entry.strip()
            if not text:
                continue
            match = re.match(r'^\[(\d+)\]\s*(.+)$', text)
            if match:
                references.append({'index': idx, 'text': match.group(2).strip(), 'raw': text})
            else:
                references.append({'index': idx, 'text': text, 'raw': text})
        return references
