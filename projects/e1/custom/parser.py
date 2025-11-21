"""
中国科学技术大学本科毕业论文格式 - 解析器

功能：解析normalized.txt文本文件，识别标题、段落、图表、公式、参考文献等元素
特点：支持该校特定的标题格式（第X章、第X节、X、等）
"""

import os
import re
from glob import glob


class E1Parser:
    """中国科学技术大学论文格式解析器"""

    def __init__(self, image_dir='projects/e1/input/images'):
        """
        初始化解析器

        Args:
            image_dir: 图片文件所在目录
        """
        self.image_dir = image_dir
        # 中文数字映射表
        self.chinese_numbers = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
            '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
            '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20
        }

    def _chinese_to_arabic(self, chinese_num):
        """
        将中文数字转换为阿拉伯数字

        Args:
            chinese_num: 中文数字字符串（如"一"、"十一"、"二十"）

        Returns:
            int: 对应的阿拉伯数字

        Examples:
            "一" -> 1
            "十" -> 10
            "十一" -> 11
            "二十" -> 20
        """
        # 直接查表
        if chinese_num in self.chinese_numbers:
            return self.chinese_numbers[chinese_num]

        # 处理"十一"到"十九"
        if chinese_num.startswith('十') and len(chinese_num) == 2:
            return 10 + self.chinese_numbers.get(chinese_num[1], 0)

        # 处理"二十"、"三十"等
        if chinese_num.endswith('十'):
            return self.chinese_numbers.get(chinese_num[0], 0) * 10

        # 处理"二十一"、"二十二"等
        if '十' in chinese_num and len(chinese_num) == 3:
            tens = self.chinese_numbers.get(chinese_num[0], 0) * 10
            ones = self.chinese_numbers.get(chinese_num[2], 0)
            return tens + ones

        # 无法识别时返回0
        return 0

    def _is_heading1(self, line):
        """识别一级标题：第X章"""
        return re.match(r'^第([一二三四五六七八九十]+)章\s+(.+)$', line)

    def _is_heading2(self, line):
        """识别二级标题：第X节"""
        return re.match(r'^第([一二三四五六七八九十]+)节\s+(.+)$', line)

    def _is_heading3(self, line):
        """识别三级标题：X、"""
        return re.match(r'^([一二三四五六七八九十]+)、\s*(.*)$', line)

    def _is_heading4(self, line):
        """识别四级标题：X."""
        return re.match(r'^(\d+)\.\s+(.+)$', line)

    def _is_heading5(self, line):
        """识别五级标题：（X）"""
        return re.match(r'^[（(](\d+)[）)]\s+(.+)$', line)

    def _is_references_header(self, line):
        """识别参考文献标题"""
        return line in ['参考文献', '参 考 文 献', '[REFERENCES]']

    def _is_acknowledgements_header(self, line):
        """识别致谢标题"""
        return line in ['致谢', '致 谢', '[ACKNOWLEDGEMENTS]', '[致谢]']

    def _is_appendix_header(self, line):
        """识别附录标题"""
        return line in ['附录', '附 录', '[APPENDIX]', '[附录]']

    def _is_new_reference_entry(self, line, has_current):
        """判断是否是新的参考文献条目"""
        # 以数字+点号开头，如 "1."
        if re.match(r'^\d+\.\s+', line):
            return True
        # 以[数字]开头，如 "[1]"
        if re.match(r'^\[\d+\]\s+', line):
            return True
        return False

    def _resolve_figure_image_path(self, figure_number, filename_hint=None):
        """
        解析图片文件路径

        Args:
            figure_number: 图片编号（如"1-1"）
            filename_hint: 文件名提示

        Returns:
            str: 图片文件的完整路径，找不到则返回None
        """
        if not os.path.exists(self.image_dir):
            return None

        # 如果有文件名提示，直接使用
        if filename_hint:
            candidate = os.path.join(self.image_dir, filename_hint)
            if os.path.exists(candidate):
                return candidate

        # 尝试匹配图片编号
        # 格式：图1-1.png, fig1-1.jpg, figure_1_1.png等
        patterns = [
            f"图{figure_number}.*",
            f"fig{figure_number.replace('-', '_')}.*",
            f"fig{figure_number.replace('-', '-')}.*",
            f"figure_{figure_number.replace('-', '_')}.*",
            f"figure{figure_number.replace('-', '')}.*",
            f"{figure_number.replace('-', '_')}.*",
            f"{figure_number.replace('-', '-')}.*"
        ]

        for pattern in patterns:
            matches = glob(os.path.join(self.image_dir, pattern))
            if matches:
                return matches[0]

        return None

    def parse_file(self, file_path):
        """
        解析文件并返回结构化内容

        Args:
            file_path: normalized.txt文件路径

        Returns:
            dict: 包含论文各部分内容的字典
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return self.parse_text(text)

    def parse_text(self, text):
        """
        解析文本内容

        Args:
            text: 预处理后的论文文本

        Returns:
            dict: 结构化的论文内容
        """
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
        current_paragraph = []
        references_buffer = []
        current_reference_lines = []

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # 处理空行
            if not line:
                if current_section == 'references':
                    # 参考文献中的空行：刷新当前条目
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                elif current_paragraph and current_section:
                    # 刷新当前段落
                    para_text = ' '.join(current_paragraph)
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
                i += 1
                continue

            # 提取论文标题（第一行非标记行）
            if i == 0 and not line.startswith('#') and not line.startswith('['):
                content['title'] = line
                i += 1
                continue

            # 识别摘要标题
            if line.startswith('[ABSTRACT]') or line == '摘要' or line == '摘 要':
                current_section = 'abstract'
                i += 1
                continue

            # 识别中文关键词
            if line.startswith('关键词：') or line.startswith('关键词:'):
                # 先刷新当前段落
                if current_paragraph and current_section == 'abstract':
                    content['abstract']['content'].append(' '.join(current_paragraph))
                    current_paragraph = []
                keywords_text = line.split('：', 1)[-1].split(':', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            # 识别英文摘要标题
            if line.upper().startswith('[ABSTRACT_EN]') or line.upper() == 'ABSTRACT':
                current_section = 'abstract_en'
                i += 1
                continue

            # 识别英文关键词
            if line.lower().startswith('key words:') or line.lower().startswith('keywords:'):
                if current_paragraph and current_section == 'abstract_en':
                    content['abstract_en']['content'].append(' '.join(current_paragraph))
                    current_paragraph = []
                keywords_text = line.split(':', 1)[-1]
                keywords = [k.strip() for k in re.split('[;；]', keywords_text) if k.strip()]
                content['abstract_en']['keywords'] = keywords
                current_section = None
                i += 1
                continue

            # 识别正文开始标记
            if line.startswith('[BODY]'):
                current_section = 'body'
                i += 1
                continue

            # 识别参考文献
            if self._is_references_header(line):
                # 刷新当前段落
                if current_paragraph and current_section == 'body' and current_chapter:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []
                current_section = 'references'
                current_reference_lines = []
                i += 1
                continue

            # 在参考文献区域内处理
            if current_section == 'references':
                if line.upper().startswith('[/REFERENCES]'):
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                        current_reference_lines = []
                    current_section = None
                    i += 1
                    continue

                # 判断是否是新条目
                if self._is_new_reference_entry(line, bool(current_reference_lines)):
                    if current_reference_lines:
                        references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = [line]
                else:
                    current_reference_lines.append(line)
                i += 1
                continue

            # 识别致谢
            if self._is_acknowledgements_header(line):
                # 刷新参考文献缓冲
                if current_reference_lines:
                    references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = []
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    if current_section == 'body' and current_chapter:
                        current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    elif current_section == 'acknowledgements':
                        content['acknowledgements'].append(para_text)
                    elif current_section == 'appendix':
                        content['appendix'].append(para_text)
                    current_paragraph = []
                current_section = 'acknowledgements'
                i += 1
                continue

            # 识别致谢结束标记
            if line.upper().startswith('[/ACKNOWLEDGEMENTS]') or line == '[/致谢]':
                if current_paragraph and current_section == 'acknowledgements':
                    content['acknowledgements'].append(' '.join(current_paragraph))
                    current_paragraph = []
                current_section = None
                i += 1
                continue

            # 识别附录
            if self._is_appendix_header(line):
                # 刷新参考文献缓冲
                if current_reference_lines:
                    references_buffer.append(' '.join(current_reference_lines).strip())
                    current_reference_lines = []
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    if current_section == 'body' and current_chapter:
                        current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    elif current_section == 'acknowledgements':
                        content['acknowledgements'].append(para_text)
                    elif current_section == 'appendix':
                        content['appendix'].append(para_text)
                    current_paragraph = []
                current_section = 'appendix'
                i += 1
                continue

            # 识别附录结束标记
            if line.upper().startswith('[/APPENDIX]') or line == '[/附录]':
                if current_paragraph and current_section == 'appendix':
                    content['appendix'].append(' '.join(current_paragraph))
                    current_paragraph = []
                current_section = None
                i += 1
                continue

            # 识别一级标题：第X章
            match_h1 = self._is_heading1(line)
            if match_h1:
                # 刷新当前段落
                if current_paragraph and current_chapter:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []

                chapter_num = self._chinese_to_arabic(match_h1.group(1))
                chapter_title = match_h1.group(2).strip()
                current_chapter = {
                    'number': chapter_num,
                    'title': chapter_title,
                    'content': []
                }
                content['chapters'].append(current_chapter)
                i += 1
                continue

            # 识别二级标题：第X节
            match_h2 = self._is_heading2(line)
            if match_h2 and current_chapter:
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []

                section_num = self._chinese_to_arabic(match_h2.group(1))
                section_title = match_h2.group(2).strip()
                current_chapter['content'].append({
                    'type': 'heading2',
                    'number': f"{current_chapter['number']}.{section_num}",
                    'text': section_title,
                    'chinese_number': match_h2.group(1)
                })
                i += 1
                continue

            # 识别三级标题：X、
            match_h3 = self._is_heading3(line)
            if match_h3 and current_chapter:
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []

                subsection_num = self._chinese_to_arabic(match_h3.group(1))
                subsection_title = match_h3.group(2).strip() if match_h3.group(2) else ''
                current_chapter['content'].append({
                    'type': 'heading3',
                    'number': subsection_num,
                    'text': subsection_title,
                    'chinese_number': match_h3.group(1)
                })
                i += 1
                continue

            # 识别四级标题：X.
            match_h4 = self._is_heading4(line)
            if match_h4 and current_chapter:
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []

                current_chapter['content'].append({
                    'type': 'heading4',
                    'number': match_h4.group(1),
                    'text': match_h4.group(2).strip()
                })
                i += 1
                continue

            # 识别五级标题：（X）
            match_h5 = self._is_heading5(line)
            if match_h5 and current_chapter:
                # 刷新当前段落
                if current_paragraph:
                    para_text = ' '.join(current_paragraph)
                    current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                    current_paragraph = []

                current_chapter['content'].append({
                    'type': 'heading5',
                    'number': match_h5.group(1),
                    'text': match_h5.group(2).strip()
                })
                i += 1
                continue

            # 识别图片标记：[FIGURE:编号]
            if line.startswith('[FIGURE:'):
                match_fig = re.match(r'\[FIGURE:([\d\-]+)\]', line)
                if match_fig and current_chapter:
                    # 刷新当前段落
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                        current_paragraph = []

                    i += 1
                    if i < len(lines):
                        caption_line = lines[i].strip()
                        # 格式：题注|来源|文件名提示
                        parts = [part.strip() for part in caption_line.split('|')]
                        caption = parts[0] if parts else ''
                        source = parts[1] if len(parts) > 1 and parts[1] else None
                        image_hint = parts[2] if len(parts) > 2 and parts[2] else None

                        figure_number = match_fig.group(1)
                        image_path = self._resolve_figure_image_path(figure_number, filename_hint=image_hint)

                        current_chapter['content'].append({
                            'type': 'figure',
                            'number': figure_number,
                            'caption': caption,
                            'source': source,
                            'path': image_path
                        })

                    i += 1
                    # 跳过结束标记
                    if i < len(lines) and lines[i].strip() == '[/FIGURE]':
                        i += 1
                    continue
                i += 1
                continue

            # 识别表格标记：[TABLE:编号]
            if line.startswith('[TABLE:'):
                match_tbl = re.match(r'\[TABLE:([\d\-]+)\]', line)
                if match_tbl and current_chapter:
                    # 刷新当前段落
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                        current_paragraph = []

                    i += 1
                    table_lines = []
                    caption = ''
                    source = None

                    if i < len(lines):
                        caption_line = lines[i].strip()
                        # 格式：题注|来源
                        parts = [part.strip() for part in caption_line.split('|')]
                        caption = parts[0] if parts else ''
                        source = parts[1] if len(parts) > 1 and parts[1] else None
                        i += 1

                    # 读取表格内容（直到[/TABLE]）
                    while i < len(lines):
                        tbl_line = lines[i].strip()
                        if tbl_line == '[/TABLE]':
                            i += 1
                            break
                        if tbl_line:  # 跳过空行
                            table_lines.append(tbl_line)
                        i += 1

                    # 解析表格数据（假设用|分隔列）
                    table_data = []
                    for tbl_line in table_lines:
                        row = [cell.strip() for cell in tbl_line.split('|')]
                        table_data.append(row)

                    current_chapter['content'].append({
                        'type': 'table',
                        'number': match_tbl.group(1),
                        'caption': caption,
                        'source': source,
                        'data': table_data
                    })
                    continue
                i += 1
                continue

            # 识别公式标记：[FORMULA:编号]
            if line.startswith('[FORMULA:'):
                match_formula = re.match(r'\[FORMULA:([\d\-]+)\]', line)
                if match_formula and current_chapter:
                    # 刷新当前段落
                    if current_paragraph:
                        para_text = ' '.join(current_paragraph)
                        current_chapter['content'].append({'type': 'paragraph', 'text': para_text})
                        current_paragraph = []

                    i += 1
                    formula_content = ''
                    if i < len(lines):
                        formula_content = lines[i].strip()
                        i += 1

                    # 跳过结束标记
                    if i < len(lines) and lines[i].strip() == '[/FORMULA]':
                        i += 1

                    current_chapter['content'].append({
                        'type': 'formula',
                        'number': match_formula.group(1),
                        'content': formula_content
                    })
                    continue
                i += 1
                continue

            # 普通文本行：累积到当前段落
            if current_section:
                current_paragraph.append(line)

            i += 1

        # 刷新最后的缓冲区
        if current_reference_lines:
            references_buffer.append(' '.join(current_reference_lines).strip())

        if current_paragraph:
            para_text = ' '.join(current_paragraph)
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

        # 将references_buffer复制到content['references']
        content['references'] = references_buffer

        return content
