import os
import sys

# 项目内模块
from custom.styles import StyleManager
from custom.parser import ContentParser
from custom.formatter import ThesisGenerator


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_dir, 'config', 'thesis_format.json')
    input_path = os.path.join(base_dir, 'input', 'normalized.txt')
    output_path = os.path.join(base_dir, 'output', 'thesis.docx')

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    try:
        style_manager = StyleManager(config_path)
    except Exception as exc:
        print(f"加载配置失败: {exc}")
        sys.exit(1)

    try:
        parser = ContentParser(image_dir=os.path.join(base_dir, 'input'))
        content = parser.parse_file(input_path)
    except Exception as exc:
        print(f"解析输入失败: {exc}")
        sys.exit(1)

    try:
        generator = ThesisGenerator(style_manager)
        generator.generate(content, output_path)
    except Exception as exc:
        print(f"生成文档失败: {exc}")
        sys.exit(1)

    print(f"生成完成: {output_path}")


if __name__ == '__main__':
    main()
