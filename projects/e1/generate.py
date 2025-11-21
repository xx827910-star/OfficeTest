"""Entry point for generating the e1 thesis docx."""
import os
from custom import E1ContentParser, E1StyleManager, ThesisGenerator

PROJECT_ID = 'e1'
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_TEXT = os.path.join(PROJECT_ROOT, 'input', 'normalized.txt')
CONFIG_PATH = os.path.join(PROJECT_ROOT, 'config', 'thesis_format.json')
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'output')
OUTPUT_PATH = os.path.join(OUTPUT_DIR, f'{PROJECT_ID}thesis.docx')


def main():
    image_dirs = [
        os.path.join(PROJECT_ROOT, 'input', 'images'),
        os.path.join(PROJECT_ROOT, '..', '..', 'examples', 'extracted', PROJECT_ID, 'images')
    ]
    parser = E1ContentParser(image_roots=image_dirs)
    content = parser.parse_file(INPUT_TEXT)

    style_manager = E1StyleManager(CONFIG_PATH)
    formatter = ThesisGenerator(style_manager)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    formatter.generate(content, OUTPUT_PATH)
    print(f'生成完成: {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
