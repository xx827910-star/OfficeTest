"""Entry point for generating the e5 thesis document."""
from __future__ import annotations

from pathlib import Path

from custom import StyleManager, ThesisFormatter, ThesisParser

PROJECT_ID = 'e5'


def main() -> None:
    project_root = Path(__file__).resolve().parent
    config_path = project_root / 'config' / 'thesis_format.json'
    normalized_path = project_root / 'input' / 'normalized.txt'
    image_dir = project_root / 'input' / 'images'
    output_dir = project_root / 'output'
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f'{PROJECT_ID}thesis.docx'

    parser = ThesisParser(image_dir=image_dir)
    content = parser.parse_file(normalized_path)

    style_manager = StyleManager(config_path)
    formatter = ThesisFormatter(style_manager)
    formatter.generate(content, str(output_path))


if __name__ == '__main__':
    main()
