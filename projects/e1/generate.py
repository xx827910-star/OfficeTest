import argparse
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(BASE_DIR)  # projects
REPO_ROOT = os.path.dirname(REPO_ROOT)  # repo root
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from custom import E1ContentParser, E1StyleManager, E1ThesisFormatter

PROJECT_ID = 'e1'
DEFAULT_CONFIG = os.path.join(BASE_DIR, 'config', 'thesis_format.json')
DEFAULT_INPUT = os.path.join(BASE_DIR, 'input', 'normalized.txt')
DEFAULT_OUTPUT = os.path.join(BASE_DIR, 'output', f'{PROJECT_ID}thesis.docx')


def parse_args():
    parser = argparse.ArgumentParser(description='Generate the formatted thesis document for project e1.')
    parser.add_argument('--config', default=DEFAULT_CONFIG, help='Path to thesis_format.json')
    parser.add_argument('--input', default=DEFAULT_INPUT, help='Path to normalized.txt content file')
    parser.add_argument('--output', default=DEFAULT_OUTPUT, help='Destination docx path')
    parser.add_argument('--images', nargs='*', default=None, help='Optional override for image search directories')
    return parser.parse_args()


def main():
    args = parse_args()
    if not os.path.exists(args.config):
        raise FileNotFoundError(f'Config not found: {args.config}')
    if not os.path.exists(args.input):
        raise FileNotFoundError(f'Input not found: {args.input}')
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    style_manager = E1StyleManager(args.config)
    parser = E1ContentParser(image_dirs=args.images)
    content = parser.parse_file(args.input)

    formatter = E1ThesisFormatter(style_manager)
    formatter.generate(content, args.output, include_toc=True)

    audit = formatter.audit_log
    missing_assets = parser.get_missing_assets()
    self_check = {
        'toc_inserted': audit.get('toc_inserted', False),
        'bookmark_count': audit.get('bookmark_count', 0),
        'cn_abstract_present': bool(content.get('abstract', {}).get('content')),
        'en_abstract_present': bool(content.get('abstract_en', {}).get('content')),
        'omml_enabled': audit.get('omml_used', False),
        'seq_fields': audit.get('seq_fields_used', False),
        'heading_eastAsia': audit.get('heading_runs_have_east_asia', False),
        'paragraph_spacing_enforced': audit.get('paragraph_spacing_applied', False),
        'missing_figures': audit.get('missing_figures', []) or missing_assets,
    }

    print('--- Self-check log ---')
    for key, value in self_check.items():
        print(f'{key}: {value}')
    print(f'Document generated at: {args.output}')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'Generation failed: {exc}', file=sys.stderr)
        sys.exit(1)
