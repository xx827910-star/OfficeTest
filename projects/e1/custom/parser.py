"""E1 项目专用内容解析器。"""
import os
import re
from glob import glob
from typing import Iterable, List, Optional


class E1ContentParser:
    """解析 normalized.txt，输出结构化章节数据。"""

    def __init__(self, image_roots: Optional[Iterable[str]] = None):
        """初始化解析器并准备图片搜索路径。"""
        roots = list(image_roots or ['examples/images'])
        normalized_roots: List[str] = []
        for root in roots:
            if not root:
                continue
            normalized_roots.append(os.path.abspath(root))
        if not normalized_roots:
            normalized_roots.append(os.path.abspath('examples/images'))
        self.image_dirs = normalized_roots

    def parse_file(self, file_path):
        """
        从文件解析内容
        :param file_path: 文件路径
        :return: 解析后的内容结构
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return self.parse_text(text)

    def parse_text(self, text):
        """以鲁棒方式解析 normalized.txt 文本。"""
        lines = text.split('\n')
        content = {
            'title': '',
            'abstract': {'content': [], 'keywords': []},
            'abstract_en': {'content': [], 'keywords': []},
            'chapters': [],
            'references': [],
            'acknowledgements': [],
            'appendix': []
        }

        current_section = None
        current_chapter = None
        current_paragraph: List[str] = []
        references_buffer: List[str] = []
        current_reference_lines: List[str] = []
        chapter_section_index = 0
        subsection_index = 0
        skip_toc_block = False

        def flush_paragraph():
            nonlocal current_paragraph
            if not current_paragraph:
                return
            para_text = ' '.join(current_paragraph).strip()
            if not para_text:
                current_paragraph = []
                return
            if current_section == 'abstract':
                content['abstract']['content'].append(para_text)
            elif current_section == 'abstract_en':
                content['abstract_en']['content'].append(para_text)
            elif current_section == 'body' and current_chapter:
                current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
            elif current_section == 'acknowledgements':
                content['acknowledgements'].append(para_text)
            elif current_section == 'appendix':
                content['appendix'].append(para_text)
            current_paragraph = []

        def flush_references_entry():
            nonlocal current_reference_lines
            if current_reference_lines:
                entry = ' '.join(current_reference_lines).strip()
                if entry:
                    references_buffer.append(entry)
                current_reference_lines = []

        i = 0
        while i < len(lines):
            raw_line = lines[i]
            line = raw_line.strip()
            upper_line = line.upper()

            if not line:
                if current_section == 'references':
                    flush_references_entry()
                else:
                    flush_paragraph()
                i += 1
                continue

            if upper_line == '[TOC]':
                skip_toc_block = True
                i += 1
                continue
            if skip_toc_block:
                if upper_line == '[/TOC]':
                    skip_toc_block = False
                i += 1
                continue

            if not content['title'] and not line.startswith('#'):
                content['title'] = line
                i += 1
                continue
            if content['title'] and content['title'].strip() and line.startswith('——'):
                content['title'] = f"{content['title']}\n{line}"
                i += 1
                continue

            if upper_line in {'[ABSTRACT]', '摘要'}:
                flush_paragraph()
                current_section = 'abstract'
                i += 1
                continue
            if upper_line in {'[/ABSTRACT]', '[/摘要]'}:
                flush_paragraph()
                current_section = None
                i += 1
                continue

            if line.startswith('关键词'):
                flush_paragraph()
                keywords_text = line.split(':', 1)[-1]
                keywords_text = keywords_text.split('：', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            if upper_line in {'[ABSTRACT_EN]', 'ABSTRACT'}:
                flush_paragraph()
                current_section = 'abstract_en'
                i += 1
                continue
            if upper_line in {'[/ABSTRACT_EN]', '[/ABSTRACT_EN ]', '[/ABSTRACT EN]'}:
                flush_paragraph()
                current_section = None
                i += 1
                continue
            if line.lower().startswith('key words') or line.lower().startswith('keywords'):
                flush_paragraph()
                keywords_text = line.split(':', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract_en']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            if upper_line == '[BODY]':
                flush_paragraph()
                current_section = 'body'
                i += 1
                continue
            if upper_line == '[/BODY]':
                flush_paragraph()
                current_section = None
                current_chapter = None
                i += 1
                continue

            if self._is_references_header(line):
                flush_paragraph()
                flush_references_entry()
                current_section = 'references'
                i += 1
                continue
            if current_section == 'references':
                if upper_line.startswith('[/REFERENCES]'):
                    flush_references_entry()
                    current_section = None
                    i += 1
                    continue
                if self._is_new_reference_entry(line, bool(current_reference_lines)):
                    flush_references_entry()
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
            if upper_line in {'[/ACKNOWLEDGEMENTS]', '[/致谢]'}:
                flush_paragraph()
                if current_section == 'acknowledgements':
                    current_section = None
                i += 1
                continue

            if self._is_appendix_header(line):
                flush_paragraph()
                current_section = 'appendix'
                i += 1
                continue
            if upper_line in {'[/APPENDIX]', '[/附录]'}:
                flush_paragraph()
                if current_section == 'appendix':
                    current_section = None
                i += 1
                continue

            if current_section != 'body' and self._match_level1(line):
                current_section = 'body'

            if current_section == 'body':
                level1 = self._match_level1(line)
                if level1:
                    flush_paragraph()
                    chapter_num, chapter_title = level1
                    current_chapter = {'number': chapter_num, 'title': chapter_title, 'content': []}
                    content['chapters'].append(current_chapter)
                    chapter_section_index = 0
                    subsection_index = 0
                    i += 1
                    continue

                if current_chapter:
                    h2 = self._match_level2(line)
                    if h2:
                        flush_paragraph()
                        chapter_section_index += 1
                        subsection_index = 0
                        current_chapter['content'].append({
                            'type': 'heading2',
                            'text': h2,
                            'section_index': chapter_section_index
                        })
                        i += 1
                        continue

                    h3 = self._match_level3(line)
                    if h3:
                        flush_paragraph()
                        subsection_index += 1
                        current_chapter['content'].append({
                            'type': 'heading3',
                            'text': h3,
                            'section_index': chapter_section_index,
                            'subsection_index': subsection_index
                        })
                        i += 1
                        continue

                    if line.startswith('[FIGURE:'):
                        flush_paragraph()
                        figure_match = re.match(r'\[FIGURE:([\d\-]+)\]', line)
                        if figure_match:
                            figure_number = figure_match.group(1)
                            i += 1
                            caption = ''
                            source = None
                            hint = None
                            if i < len(lines):
                                caption_line = lines[i].strip()
                                parts = [part.strip() for part in caption_line.split('|')]
                                caption = parts[0] if parts else ''
                                if len(parts) > 1:
                                    source = parts[1] or None
                                if len(parts) > 2:
                                    hint = parts[2] or None
                            image_path = self._resolve_figure_image_path(figure_number, filename_hint=hint)
                            current_chapter['content'].append({
                                'type': 'figure',
                                'number': figure_number,
                                'caption': caption,
                                'source': source,
                                'path': image_path,
                                'file_hint': hint
                            })
                            while i < len(lines) and lines[i].strip() != '[/FIGURE]':
                                i += 1
                            if i < len(lines) and lines[i].strip() == '[/FIGURE]':
                                i += 1
                            continue
                        i += 1
                        continue

                    if line.startswith('[TABLE:'):
                        flush_paragraph()
                        table_match = re.match(r'\[TABLE:([\d\-]+)\]', line)
                        if table_match:
                            table_number = table_match.group(1)
                            i += 1
                            caption = ''
                            source = None
                            if i < len(lines):
                                caption_line = lines[i].strip()
                                parts = [part.strip() for part in caption_line.split('|')]
                                caption = parts[0] if parts else ''
                                if len(parts) > 1:
                                    source = parts[1] or None
                                i += 1
                            rows: List[List[str]] = []
                            while i < len(lines):
                                tbl_line = lines[i].strip()
                                if tbl_line == '[/TABLE]':
                                    break
                                if tbl_line:
                                    rows.append([cell.strip() for cell in tbl_line.split('|')])
                                i += 1
                            current_chapter['content'].append({
                                'type': 'table',
                                'number': table_number,
                                'caption': caption,
                                'source': source,
                                'rows': rows
                            })
                            if i < len(lines) and lines[i].strip() == '[/TABLE]':
                                i += 1
                            continue
                        i += 1
                        continue

                    if line.startswith('[FORMULA:'):
                        flush_paragraph()
                        formula_match = re.match(r'\[FORMULA:([\d\-]+)\]', line)
                        if formula_match:
                            formula_number = formula_match.group(1)
                            i += 1
                            formula_lines: List[str] = []
                            while i < len(lines) and lines[i].strip() != '[/FORMULA]':
                                if lines[i].strip():
                                    formula_lines.append(lines[i].strip())
                                i += 1
                            current_chapter['content'].append({
                                'type': 'formula',
                                'number': formula_number,
                                'content': '\n'.join(formula_lines)
                            })
                            if i < len(lines) and lines[i].strip() == '[/FORMULA]':
                                i += 1
                            continue
                        i += 1
                        continue

            if current_section in {'abstract', 'abstract_en', 'acknowledgements', 'appendix'}:
                current_paragraph.append(line)
            elif current_section == 'body' and current_chapter:
                current_paragraph.append(line)

            i += 1

        flush_paragraph()
        flush_references_entry()
        content['references'] = self._build_references_list(references_buffer)
        return content

    def _match_level1(self, line):
        match_standard = re.match(r'^第([一二三四五六七八九十百千万\d]+)章\s*(.+)$', line)
        if match_standard:
            return self._chinese_to_arabic(match_standard.group(1)), match_standard.group(2).strip()

        match_chinese = re.match(r'^([一二三四五六七八九十百千万]+)[、\.．]\s*(.+)$', line)
        if match_chinese:
            return self._chinese_to_arabic(match_chinese.group(1)), match_chinese.group(2).strip()

        match_numeric = re.match(r'^(\d+)\s+(.+)$', line)
        if match_numeric:
            return int(match_numeric.group(1)), match_numeric.group(2).strip()
        return None

    def _match_level2(self, line):
        match_parenthetical = re.match(r'^（([一二三四五六七八九十百千万\d]+)）\s*(.+)$', line)
        if match_parenthetical:
            return match_parenthetical.group(2).strip()

        match_section = re.match(r'^第([一二三四五六七八九十百千万\d]+)节\s*(.+)$', line)
        if match_section:
            return match_section.group(2).strip()

        match_decimal = re.match(r'^(\d+\.\d+)\s+(.+)$', line)
        if match_decimal:
            return match_decimal.group(2).strip()
        return None

    def _match_level3(self, line):
        match_decimal = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)$', line)
        if match_decimal:
            return match_decimal.group(2).strip()

        match_simple = re.match(r'^(\d+)[\.、．]\s*(.+)$', line)
        if match_simple:
            return match_simple.group(2).strip()
        return None

    def _resolve_figure_image_path(self, figure_number, filename_hint=None):
        """根据图号或额外提示解析图片路径"""
        search_candidates = []

        if filename_hint:
            hint_path = filename_hint
            if not os.path.isabs(hint_path):
                for root in self.image_dirs:
                    search_candidates.append(os.path.join(root, hint_path))
            else:
                search_candidates.append(hint_path)

        normalized_num = figure_number.replace('-', '_')
        base_name = f'figure_{normalized_num}'
        extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff']

        for root in self.image_dirs:
            for ext in extensions:
                search_candidates.append(os.path.join(root, f'{base_name}{ext}'))
                search_candidates.append(os.path.join(root, f'{base_name}{ext.upper()}'))

        seen = set()
        for candidate in search_candidates:
            if not candidate or candidate in seen:
                continue
            seen.add(candidate)
            if os.path.exists(candidate):
                return candidate

        for root in self.image_dirs:
            pattern = os.path.join(root, f'{base_name}.*')
            matches = sorted(glob(pattern))
            if matches:
                return matches[0]
        return None

    def _chinese_to_arabic(self, chinese_num):
        """
        转换中文数字到阿拉伯数字
        :param chinese_num: 中文数字（如"一"、"二"）
        :return: 阿拉伯数字
        """
        chinese_map = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
        }

        # 如果已经是数字，直接返回
        if chinese_num.isdigit():
            return int(chinese_num)

        # 简单的中文数字转换（支持一到十）
        if chinese_num in chinese_map:
            return chinese_map[chinese_num]

        # 处理十几、几十等
        if '十' in chinese_num:
            if chinese_num == '十':
                return 10
            elif chinese_num.startswith('十'):
                return 10 + chinese_map.get(chinese_num[1], 0)
            elif chinese_num.endswith('十'):
                return chinese_map.get(chinese_num[0], 0) * 10
            else:
                parts = chinese_num.split('十')
                return chinese_map.get(parts[0], 0) * 10 + chinese_map.get(parts[1], 0)

        return 1  # 默认返回1

    def _is_references_header(self, line):
        """
        判断是否为参考文献区块的标题
        """
        normalized = line.strip()
        reference_headers = {
            '[REFERENCES]', 'REFERENCES', '参考文献', '参考文献：', '［参考文献］',
            '[参考文献]', '【参考文献】', '参考文献:', 'References', 'REFERENCES：'
        }
        if normalized in reference_headers:
            return True
        normalized_plain = normalized.strip('[]［］【】():：')
        return normalized_plain.upper() == 'REFERENCES' or normalized_plain == '参考文献'

    def _is_acknowledgements_header(self, line):
        """
        判断是否为致谢部分标题
        """
        normalized = line.strip()
        ack_headers = {
            '[ACKNOWLEDGEMENTS]', 'ACKNOWLEDGEMENTS', '致谢', '致  谢', '致 谢', '[致谢]', '【致谢】'
        }
        if normalized in ack_headers:
            return True
        normalized_plain = normalized.strip('[]［］【】():：')
        return normalized_plain.upper() == 'ACKNOWLEDGEMENTS' or normalized_plain == '致谢'

    def _is_appendix_header(self, line):
        """
        判断是否为附录部分标题
        """
        normalized = line.strip()
        appendix_headers = {
            '[APPENDIX]', 'APPENDIX', '附录', '附  录', '[附录]', '【附录】'
        }
        if normalized in appendix_headers:
            return True
        normalized_plain = normalized.strip('[]［］【】():：')
        return normalized_plain.upper() == 'APPENDIX' or normalized_plain == '附录'

    def _is_new_reference_entry(self, line, has_current_entry):
        """
        判断当前行是否代表新的参考文献条目
        """
        if not has_current_entry:
            return True
        return bool(re.match(r'^\[\d+\]', line))

    def _build_references_list(self, references_buffer):
        """
        将参考文献原始文本转换为结构化列表
        """
        references = []
        for idx, entry in enumerate(references_buffer, 1):
            text = entry.strip()
            if not text:
                continue
            match = re.match(r'^\[(\d+)\]\s*(.+)$', text)
            if match:
                normalized_text = match.group(2).strip()
                original_index = int(match.group(1))
            else:
                normalized_text = text
                original_index = None
            references.append({
                'index': idx,
                'original_index': original_index,
                'text': normalized_text,
                'raw': text
            })
        return references
