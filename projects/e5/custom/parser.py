"""
内容解析器 - 解析输入的文本内容，识别不同部分
"""
import os
import re
from glob import glob


class USTCContentParser:
    """解析论文内容结构"""

    def __init__(self, image_dir='examples/images'):
        """
        初始化解析器
        :param image_dir: 图片目录路径
        """
        self.content = {}
        self.image_dir = image_dir

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
        """
        解析文本内容
        :param text: 原始文本
        :return: 解析后的内容结构
        """
        lines = text.split('\n')
        content = {
            'title': '',
            'abstract': {
                'content': [],
                'keywords': []
            },
            'abstract_en': {
                'content': [],
                'keywords': []
            },
            'chapters': [],
            'references': [],
            'acknowledgements': [],
            'appendix': []
        }

        current_section = None
        current_chapter = None
        current_paragraph = []
        references_buffer = []
        current_reference_lines = []

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 空行处理
            if not line:
                if current_section == 'references':
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                elif current_paragraph and current_section:
                    para_text = ' '.join(current_paragraph)
                    if current_section == 'abstract':
                        content['abstract']['content'].append(para_text)
                    elif current_section == 'abstract_en':
                        content['abstract_en']['content'].append(para_text)
                    elif current_section == 'body' and current_chapter:
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                    elif current_section == 'acknowledgements':
                        content['acknowledgements'].append(para_text)
                    elif current_section == 'appendix':
                        content['appendix'].append(para_text)
                    current_paragraph = []
                i += 1
                continue

            # 识别标题（论文题目）
            if i == 0 and not line.startswith('#'):
                content['title'] = line
                i += 1
                continue

            # 识别摘要标记
            if line.startswith('[ABSTRACT]') or line == '摘要':
                current_section = 'abstract'
                i += 1
                continue

            # 识别中文关键词
            if line.startswith('关键词：') or line.startswith('关键词:'):
                keywords_text = line.split('：', 1)[-1].split(':', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            # 识别英文摘要标记
            if line.upper().startswith('[ABSTRACT_EN]') or line.upper() == 'ABSTRACT':
                current_section = 'abstract_en'
                i += 1
                continue

            # 识别英文关键词
            if line.lower().startswith('key words:') or line.lower().startswith('keywords:'):
                keywords_text = line.split(':', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract_en']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            # 识别正文开始
            if line.startswith('[BODY]') or re.match(r'^第[一二三四五六七八九十\d]+章', line):
                current_section = 'body'
                if re.match(r'^第[一二三四五六七八九十\d]+章', line):
                    # 已经是章节标题，不需要跳过
                    pass
                else:
                    i += 1
                    continue

            # 识别参考文献开始
            if self._is_references_header(line):
                if current_paragraph and current_section == 'body' and current_chapter:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({
                        'type': 'paragraph',
                        'text': para_text
                    })
                    current_paragraph = []
                current_section = 'references'
                current_reference_lines = []
                i += 1
                continue

            if current_section == 'references':
                if line.upper().startswith('[/REFERENCES]'):
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

            # 识别致谢部分
            if self._is_acknowledgements_header(line):
                if current_reference_lines:
                    references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = []
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    if current_section == 'body' and current_chapter:
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                    elif current_section == 'acknowledgements':
                        content['acknowledgements'].append(para_text)
                    elif current_section == 'appendix':
                        content['appendix'].append(para_text)
                    current_paragraph = []
                current_section = 'acknowledgements'
                i += 1
                continue

            if line.upper().startswith('[/ACKNOWLEDGEMENTS]') or line == '[/致谢]':
                if current_paragraph and current_section == 'acknowledgements':
                    content['acknowledgements'].append(' '.join(current_paragraph))
                    current_paragraph = []
                current_section = None
                i += 1
                continue

            # 识别附录部分
            if self._is_appendix_header(line):
                if current_reference_lines:
                    references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = []
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    if current_section == 'body' and current_chapter:
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                    elif current_section == 'acknowledgements':
                        content['acknowledgements'].append(para_text)
                    elif current_section == 'appendix':
                        content['appendix'].append(para_text)
                    current_paragraph = []
                current_section = 'appendix'
                i += 1
                continue

            if line.upper().startswith('[/APPENDIX]') or line == '[/附录]':
                if current_paragraph and current_section == 'appendix':
                    content['appendix'].append(' '.join(current_paragraph))
                    current_paragraph = []
                current_section = None
                i += 1
                continue

            # 识别一级标题（章）
            match1 = re.match(r'^第([一二三四五六七八九十\d]+)章\s+(.+)$', line)
            match2 = re.match(r'^(\d+)\s+(.+)$', line)
            match3 = re.match(r'^#+\s+(.+)$', line)  # Markdown 格式

            if match1 or (match2 and current_section == 'body'):
                if current_paragraph and current_chapter:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({
                        'type': 'paragraph',
                        'text': para_text
                    })
                    current_paragraph = []

                if match1:
                    chapter_num = self._chinese_to_arabic(match1.group(1))
                    chapter_title = match1.group(2).strip()
                elif match2:
                    chapter_num = int(match2.group(1))
                    chapter_title = match2.group(2).strip()

                current_chapter = {
                    'number': chapter_num,
                    'title': chapter_title,
                    'content': []
                }
                content['chapters'].append(current_chapter)
                i += 1
                continue

            # 识别二级标题
            match_h2 = re.match(r'^(\d+\.\d+)\s+(.+)$', line)
            if match_h2 and current_chapter:
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({
                        'type': 'paragraph',
                        'text': para_text
                    })
                    current_paragraph = []

                current_chapter['content'].append({
                    'type': 'heading2',
                    'number': match_h2.group(1),
                    'text': match_h2.group(2).strip()
                })
                i += 1
                continue

            # 识别三级标题
            match_h3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)$', line)
            if match_h3 and current_chapter:
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({
                        'type': 'paragraph',
                        'text': para_text
                    })
                    current_paragraph = []

                current_chapter['content'].append({
                    'type': 'heading3',
                    'number': match_h3.group(1),
                    'text': match_h3.group(2).strip()
                })
                i += 1
                continue

            # 识别图片标记 [FIGURE:1-1]
            if line.startswith('[FIGURE:'):
                match_fig = re.match(r'\[FIGURE:([\d\-]+)\]', line)
                if match_fig and current_chapter:
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                        current_paragraph = []

                    # 读取下一行获取标题和来源
                    i += 1
                    if i < len(lines):
                        caption_line = lines[i].strip()
                        parts = [part.strip() for part in caption_line.split('|')]
                        caption = parts[0] if parts else ''
                        source = parts[1] if len(parts) > 1 and parts[1] else None
                        image_hint = parts[2] if len(parts) > 2 and parts[2] else None

                        # 自动映射图片路径，支持多种扩展名及自定义文件名
                        figure_number = match_fig.group(1)
                        image_path = self._resolve_figure_image_path(
                            figure_number,
                            filename_hint=image_hint
                        )

                        current_chapter['content'].append({
                            'type': 'figure',
                            'number': match_fig.group(1),
                            'caption': caption,
                            'source': source,
                            'path': image_path
                        })

                    # 跳过 [/FIGURE]
                    i += 1
                    if i < len(lines) and lines[i].strip() == '[/FIGURE]':
                        i += 1
                    continue
                i += 1
                continue

            # 识别表格标记 [TABLE:1-1]
            if line.startswith('[TABLE:'):
                match_tbl = re.match(r'\[TABLE:([\d\-]+)\]', line)
                if match_tbl and current_chapter:
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                        current_paragraph = []

                    # 读取表格内容
                    i += 1
                    table_lines = []
                    caption = ''
                    source = None

                    # 第一行是标题|来源
                    if i < len(lines):
                        caption_line = lines[i].strip()
                        parts = caption_line.split('|')
                        caption = parts[0] if len(parts) > 0 else ''
                        source = parts[1] if len(parts) > 1 else None
                        i += 1

                    # 读取表格数据行
                    while i < len(lines):
                        tbl_line = lines[i].strip()
                        if tbl_line == '[/TABLE]':
                            break
                        if tbl_line and not tbl_line.startswith('['):
                            table_lines.append(tbl_line)
                        i += 1

                    # 解析表格数据
                    rows = []
                    for tbl_line in table_lines:
                        cells = [cell.strip() for cell in tbl_line.split('|')]
                        rows.append(cells)

                    current_chapter['content'].append({
                        'type': 'table',
                        'number': match_tbl.group(1),
                        'caption': caption,
                        'source': source,
                        'rows': rows
                    })
                i += 1
                continue

            # 识别公式标记 [FORMULA:2-1]
            if line.startswith('[FORMULA:'):
                match_formula = re.match(r'\[FORMULA:([\d\-]+)\]', line)
                if match_formula and current_chapter:
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({
                            'type': 'paragraph',
                            'text': para_text
                        })
                        current_paragraph = []

                    # 读取公式内容
                    i += 1
                    formula_lines = []
                    while i < len(lines):
                        formula_line = lines[i].strip()
                        if formula_line == '[/FORMULA]':
                            break
                        if formula_line:
                            formula_lines.append(formula_line)
                        i += 1

                    current_chapter['content'].append({
                        'type': 'formula',
                        'number': match_formula.group(1),
                        'content': '\n'.join(formula_lines)
                    })
                i += 1
                continue

            # 普通文本行
            if current_section == 'abstract':
                current_paragraph.append(line)
            elif current_section == 'abstract_en':
                current_paragraph.append(line)
            elif current_section == 'body' and current_chapter:
                current_paragraph.append(line)
            elif current_section == 'acknowledgements':
                current_paragraph.append(line)
            elif current_section == 'appendix':
                current_paragraph.append(line)

            i += 1

        # 处理最后的段落
        if current_paragraph:
            para_text = ' '.join(current_paragraph)
            if current_section == 'abstract':
                content['abstract']['content'].append(para_text)
            elif current_section == 'abstract_en':
                content['abstract_en']['content'].append(para_text)
            elif current_section == 'body' and current_chapter:
                current_chapter['content'].append({
                    'type': 'paragraph',
                    'text': para_text
                })
            elif current_section == 'acknowledgements':
                content['acknowledgements'].append(para_text)
            elif current_section == 'appendix':
                content['appendix'].append(para_text)

        # 处理参考文献残留
        if current_reference_lines:
            references_buffer.append(' '.join(current_reference_lines).strip())

        content['references'] = self._build_references_list(references_buffer)

        return content

    def _resolve_figure_image_path(self, figure_number, filename_hint=None):
        """根据图号或额外提示解析图片路径"""
        search_candidates = []

        if filename_hint:
            hint_path = filename_hint
            if not os.path.isabs(hint_path):
                hint_path = os.path.join(self.image_dir, hint_path)
            search_candidates.append(hint_path)

        normalized_num = figure_number.replace('-', '_')
        base_name = f'figure_{normalized_num}'
        extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tif', '.tiff']

        for ext in extensions:
            search_candidates.append(os.path.join(self.image_dir, f'{base_name}{ext}'))
            search_candidates.append(os.path.join(self.image_dir, f'{base_name}{ext.upper()}'))

        seen = set()
        for candidate in search_candidates:
            if not candidate or candidate in seen:
                continue
            seen.add(candidate)
            if os.path.exists(candidate):
                return candidate

        pattern = os.path.join(self.image_dir, f'{base_name}.*')
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
