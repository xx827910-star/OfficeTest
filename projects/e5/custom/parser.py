"""Parser for normalized thesis text (project e5)."""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple


class ThesisParser:
    """Parse normalized.txt into a structured content dictionary."""

    HEADING_CN_PATTERN = re.compile(r'^第([一二三四五六七八九十百零〇\d]+)章\s*(.+)$')
    HEADING_DECIMAL_PATTERN = re.compile(r'^(\d+(?:\.\d+){0,3})\s+(.+)$')
    HEADING_ROMAN_PATTERN = re.compile(r'^(?:Chapter|CHAPTER)\s+(\d+)\.?\s*(.+)$', re.IGNORECASE)
    CN_KEYWORDS_PATTERN = re.compile(r'^关\s*键\s*词[：:]?(.*)$')
    EN_KEYWORDS_PATTERN = re.compile(r'^(?:key\s*words?|KEY\s*WORDS?)\s*[:：](.*)$', re.IGNORECASE)

    def __init__(self, image_dir: Optional[Path] = None):
        self.image_dir = Path(image_dir) if image_dir else None
        self._image_candidates = self._discover_images()

    def parse_file(self, file_path: Path) -> Dict:
        """Parse a file path."""
        file_path = Path(file_path)
        text = file_path.read_text(encoding='utf-8')
        return self.parse_text(text)

    def parse_text(self, text: str) -> Dict:
        """Parse normalized thesis text into structured data."""
        lines = text.splitlines()
        content = {
            'title': '',
            'abstract': {'content': [], 'keywords': []},
            'abstract_en': {'content': [], 'keywords': []},
            'chapters': [],
            'references': [],
            'acknowledgements': [],
            'appendix': []
        }

        current_section: Optional[str] = None
        current_paragraph: List[str] = []
        current_chapter: Optional[Dict] = None
        figure_counters = defaultdict(int)
        table_counters = defaultdict(int)
        formula_counters = defaultdict(int)
        references_buffer: List[str] = []
        current_reference_lines: List[str] = []
        chapter_auto_number = 0
        title_locked = False

        i = 0
        total_lines = len(lines)
        while i < total_lines:
            raw_line = lines[i]
            line = raw_line.strip()
            upper_line = line.upper()

            # Title detection (before first section marker)
            if not title_locked:
                if not line:
                    i += 1
                    continue
                if line in {'论文主标题', '题目', 'TITLE'}:
                    i += 1
                    continue
                if line.startswith('['):
                    title_locked = True
                else:
                    content['title'] = line
                    title_locked = True
                    i += 1
                    continue

            # Block markers -------------------------------------------------
            if upper_line == '[ABSTRACT]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'abstract'
                i += 1
                continue

            if upper_line == '[ABSTRACT_EN]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'abstract_en'
                i += 1
                continue

            if upper_line == '[TOC]':
                i = self._skip_block(lines, i, 'TOC')
                continue

            if upper_line == '[BODY]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'body'
                i += 1
                continue

            if upper_line == '[REFERENCES]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'references'
                references_buffer = []
                current_reference_lines = []
                i += 1
                continue

            if upper_line == '[ACKNOWLEDGEMENTS]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'acknowledgements'
                i += 1
                continue

            if upper_line == '[APPENDIX]':
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                current_section = 'appendix'
                i += 1
                continue

            if upper_line.startswith('[FIGURE:'):
                block_lines, new_index = self._collect_block(lines, i, 'FIGURE')
                media_id = self._extract_block_index(line)
                if current_chapter:
                    self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                    figure_entry = self._create_figure_entry(
                        media_id,
                        block_lines,
                        current_chapter,
                        figure_counters
                    )
                    if figure_entry:
                        current_chapter['content'].append(figure_entry)
                i = new_index
                continue

            if upper_line.startswith('[TABLE:'):
                block_lines, new_index = self._collect_block(lines, i, 'TABLE')
                media_id = self._extract_block_index(line)
                if current_chapter:
                    self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                    table_entry = self._create_table_entry(
                        media_id,
                        block_lines,
                        current_chapter,
                        table_counters
                    )
                    if table_entry:
                        current_chapter['content'].append(table_entry)
                i = new_index
                continue

            if upper_line.startswith('[FORMULA:'):
                block_lines, new_index = self._collect_block(lines, i, 'FORMULA')
                media_id = self._extract_block_index(line)
                if current_chapter:
                    self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                    formula_entry = self._create_formula_entry(
                        media_id,
                        block_lines,
                        current_chapter,
                        formula_counters
                    )
                    if formula_entry:
                        current_chapter['content'].append(formula_entry)
                i = new_index
                continue

            if upper_line in {'[/REFERENCES]', '[/ACKNOWLEDGEMENTS]', '[/APPENDIX]', '[/FIGURE]', '[/TABLE]', '[/FORMULA]'}:
                self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                if upper_line == '[/REFERENCES]':
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                    for entry in references_buffer:
                        if entry:
                            content['references'].append({'text': entry})
                    references_buffer = []
                    current_section = None
                elif upper_line == '[/ACKNOWLEDGEMENTS]':
                    current_section = None
                elif upper_line == '[/APPENDIX]':
                    current_section = None
                i += 1
                continue

            # Blank line flush
            if not line:
                if current_section == 'references':
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                else:
                    self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                i += 1
                continue

            # Section-specific parsing -------------------------------------
            if current_section == 'references':
                if self._looks_like_reference_header(line, bool(current_reference_lines)):
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = [line]
                else:
                    current_reference_lines.append(line)
                i += 1
                continue

            if self.CN_KEYWORDS_PATTERN.match(line) and current_section == 'abstract':
                content['abstract']['keywords'] = self._parse_keywords_line(line)
                current_section = None
                current_paragraph.clear()
                i += 1
                continue

            if self.EN_KEYWORDS_PATTERN.match(line) and current_section == 'abstract_en':
                content['abstract_en']['keywords'] = self._parse_keywords_line(line)
                current_section = None
                current_paragraph.clear()
                i += 1
                continue

            if current_section == 'body':
                heading_data = self._parse_heading(line)
                if heading_data:
                    self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
                    level, heading_number, heading_text = heading_data
                    if level == 1:
                        if heading_number is None:
                            chapter_auto_number += 1
                            heading_number = str(chapter_auto_number)
                        else:
                            try:
                                chapter_auto_number = int(float(heading_number))
                            except ValueError:
                                chapter_auto_number += 1
                                heading_number = str(chapter_auto_number)
                        current_chapter = {
                            'number': heading_number,
                            'title': heading_text,
                            'content': []
                        }
                        content['chapters'].append(current_chapter)
                        figure_counters[heading_number] = 0
                        table_counters[heading_number] = 0
                        formula_counters[heading_number] = 0
                    elif current_chapter:
                        entry = {
                            'type': f'heading{level}',
                            'number': heading_number,
                            'text': heading_text
                        }
                        current_chapter['content'].append(entry)
                    i += 1
                    continue

            # Default: accumulate paragraph text
            current_paragraph.append(line)
            i += 1

        # Flush remainder
        self._flush_paragraph(current_section, current_paragraph, content, current_chapter)
        if current_section == 'references' and current_reference_lines:
            references_buffer.append(' '.join(current_reference_lines).strip())
        if current_section == 'references' and references_buffer:
            for entry in references_buffer:
                if entry:
                    content['references'].append({'text': entry})

        return content

    # ------------------------------------------------------------------
    # Helper methods
    # ------------------------------------------------------------------
    def _discover_images(self) -> List[Path]:
        if not self.image_dir or not self.image_dir.exists():
            return []
        return sorted(self.image_dir.glob('*.png')) + sorted(self.image_dir.glob('*.jpg'))

    def _resolve_image_path(self, media_index: Optional[int]) -> Optional[str]:
        if media_index is None or not self._image_candidates:
            return None
        idx = max(0, media_index - 1)
        if idx < len(self._image_candidates):
            return str(self._image_candidates[idx])
        return None

    def _collect_block(self, lines: List[str], start_index: int, tag: str) -> Tuple[List[str], int]:
        closing = f'[/{tag.upper()}]'
        block: List[str] = []
        i = start_index + 1
        while i < len(lines):
            current = lines[i].strip()
            if current.upper() == closing:
                return block, i + 1
            block.append(lines[i].rstrip('\n'))
            i += 1
        return block, i

    def _skip_block(self, lines: List[str], start_index: int, tag: str) -> int:
        closing = f'[/{tag.upper()}]'
        i = start_index + 1
        while i < len(lines):
            if lines[i].strip().upper() == closing:
                return i + 1
            i += 1
        return i

    def _extract_block_index(self, header_line: str) -> Optional[int]:
        match = re.search(r':(\d+)', header_line)
        if match:
            return int(match.group(1))
        return None

    def _split_caption_and_meta(self, line: str) -> Tuple[str, Dict[str, str]]:
        parts = [seg.strip() for seg in line.split('|') if seg.strip()]
        if not parts:
            return '', {}
        caption = parts[0]
        metadata: Dict[str, str] = {}
        for part in parts[1:]:
            if ':' in part:
                key, value = part.split(':', 1)
            elif '：' in part:
                key, value = part.split('：', 1)
            else:
                key, value = '备注', part
            metadata[key.strip()] = value.strip()
        return caption, metadata

    def _normalize_caption(self, caption: str) -> str:
        cleaned = caption.strip().strip('|')
        cleaned = re.sub(r'^图[一二三四五六七八九十百零〇\d]+[：:\s]+', '', cleaned)
        cleaned = re.sub(r'^表[一二三四五六七八九十百零〇\d]+[：:\s]+', '', cleaned)
        return cleaned.strip()

    def _create_figure_entry(
        self,
        media_index: Optional[int],
        block_lines: List[str],
        chapter: Dict,
        figure_counters: defaultdict
    ) -> Optional[Dict]:
        if not block_lines:
            return None
        caption_line, metadata = self._split_caption_and_meta(block_lines[0])
        caption = self._normalize_caption(caption_line)
        chapter_num = chapter.get('number', '1')
        figure_counters[chapter_num] += 1
        number = f"{chapter_num}-{figure_counters[chapter_num]}"
        source = metadata.get('来源') or metadata.get('source') or metadata.get('备注')
        return {
            'type': 'figure',
            'chapter': chapter_num,
            'number': number,
            'identifier': media_index,
            'caption': caption,
            'path': self._resolve_image_path(media_index),
            'source': source
        }

    def _create_table_entry(
        self,
        media_index: Optional[int],
        block_lines: List[str],
        chapter: Dict,
        table_counters: defaultdict
    ) -> Optional[Dict]:
        if not block_lines:
            return None
        caption_line, metadata = self._split_caption_and_meta(block_lines[0])
        caption = self._normalize_caption(caption_line)
        data_lines = [line.strip('|') for line in block_lines[1:] if line.strip()]
        rows = [
            [cell.strip() for cell in row.split('|')]
            for row in data_lines if row
        ]
        if not rows:
            return None
        chapter_num = chapter.get('number', '1')
        table_counters[chapter_num] += 1
        number = f"{chapter_num}-{table_counters[chapter_num]}"
        source = metadata.get('来源') or metadata.get('source') or metadata.get('备注')
        return {
            'type': 'table',
            'chapter': chapter_num,
            'number': number,
            'identifier': media_index,
            'caption': caption,
            'rows': rows,
            'source': source
        }

    def _create_formula_entry(
        self,
        media_index: Optional[int],
        block_lines: List[str],
        chapter: Dict,
        formula_counters: defaultdict
    ) -> Optional[Dict]:
        formula_text = '\n'.join(line.strip() for line in block_lines if line.strip())
        if not formula_text:
            return None
        chapter_num = chapter.get('number', '1')
        formula_counters[chapter_num] += 1
        number = f"{chapter_num}-{formula_counters[chapter_num]}"
        return {
            'type': 'formula',
            'chapter': chapter_num,
            'number': number,
            'identifier': media_index,
            'content': formula_text
        }

    def _flush_paragraph(
        self,
        section: Optional[str],
        paragraph_buffer: List[str],
        content: Dict,
        current_chapter: Optional[Dict]
    ) -> None:
        if not paragraph_buffer:
            return
        text = ' '.join(part.strip() for part in paragraph_buffer).strip()
        paragraph_buffer.clear()
        if not text:
            return
        if section == 'abstract':
            content['abstract']['content'].append(text)
        elif section == 'abstract_en':
            content['abstract_en']['content'].append(text)
        elif section == 'body' and current_chapter:
            current_chapter['content'].append({'type': 'paragraph', 'text': text})
        elif section == 'acknowledgements':
            content['acknowledgements'].append(text)
        elif section == 'appendix':
            content['appendix'].append(text)

    def _parse_keywords_line(self, line: str) -> List[str]:
        if '：' in line:
            _, payload = line.split('：', 1)
        elif ':' in line:
            _, payload = line.split(':', 1)
        else:
            payload = line
        parts = re.split(r'[;；、]', payload)
        return [part.strip().capitalize() for part in parts if part.strip()]

    def _parse_heading(self, line: str) -> Optional[Tuple[int, str, str]]:
        match = self.HEADING_CN_PATTERN.match(line)
        if match:
            number_text = match.group(1)
            heading_number = str(self._chinese_to_arabic(number_text))
            heading_text = match.group(2).strip()
            return 1, heading_number, heading_text

        match = self.HEADING_DECIMAL_PATTERN.match(line)
        if match:
            number_text = match.group(1)
            heading_text = match.group(2).strip()
            level = number_text.count('.') + 1
            return level, number_text, heading_text

        match = self.HEADING_ROMAN_PATTERN.match(line)
        if match:
            heading_number = match.group(1)
            heading_text = match.group(2).strip()
            return 1, heading_number, heading_text

        if line in {'结语', '总结', 'Conclusion'}:
            return 1, None, line

        return None

    def _looks_like_reference_header(self, line: str, has_buffer: bool) -> bool:
        return bool(re.match(r'^\[\d+\]', line) or (not has_buffer and line.startswith('[')))

    def _chinese_to_arabic(self, text: str) -> int:
        try:
            return int(text)
        except ValueError:
            pass

        digits = {
            '零': 0, '〇': 0,
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9
        }
        units = {'十': 10, '百': 100, '千': 1000}

        result = 0
        temp = 0
        for char in text:
            if char in digits:
                temp = digits[char]
            elif char in units:
                unit_value = units[char]
                if temp == 0:
                    temp = 1
                result += temp * unit_value
                temp = 0
            elif char.isdigit():
                temp = temp * 10 + int(char)
        result += temp
        return result or 1
