"""项目 e5 的文档生成入口脚本。"""
from pathlib import Path
import sys

from custom import USTCContentParser, USTCStyleManager, USTCFormatter


def main():
    base_dir = Path(__file__).resolve().parent
    config_path = base_dir / 'config' / 'thesis_format.json'
    normalized_path = base_dir / 'input' / 'normalized.txt'
    image_dir = base_dir / 'input' / 'images'
    output_dir = base_dir / 'output'
    output_path = output_dir / 'e5thesis.docx'

    if not config_path.exists():
        raise FileNotFoundError(f'缺少配置文件: {config_path}')
    if not normalized_path.exists():
        raise FileNotFoundError(f'缺少标准化文本: {normalized_path}')

    output_dir.mkdir(parents=True, exist_ok=True)

    style_manager = USTCStyleManager(str(config_path))
    parser = USTCContentParser(image_dir=str(image_dir))
    content = parser.parse_file(str(normalized_path))

    formatter = USTCFormatter(style_manager)
    formatter.generate(content, str(output_path))
    print(f'✓ 已生成文档: {output_path}')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:  # pragma: no cover - 入口脚本告警
        print(f'✗ 生成失败: {exc}')
        sys.exit(1)
