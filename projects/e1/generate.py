"""
中国科学技术大学本科毕业论文格式 - 文档生成主程序

用法：
    python projects/e1/generate.py

输入文件：
    - projects/e1/input/normalized.txt: 预处理后的论文文本
    - projects/e1/config/thesis_format.json: 格式配置文件

输出文件：
    - projects/e1/output/e1thesis.docx: 生成的Word文档
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from projects.e1.custom import E1Parser, E1Formatter, E1StyleManager


def main():
    """主函数"""
    # 定义路径
    base_dir = Path(__file__).parent
    input_file = base_dir / 'input' / 'normalized.txt'
    config_file = base_dir / 'config' / 'thesis_format.json'
    output_dir = base_dir / 'output'
    output_file = output_dir / 'e1thesis.docx'

    # 检查输入文件
    if not input_file.exists():
        print(f"错误：输入文件不存在: {input_file}")
        return 1

    if not config_file.exists():
        print(f"错误：配置文件不存在: {config_file}")
        return 1

    # 创建输出目录
    output_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("中国科学技术大学本科毕业论文格式 - 文档生成器")
    print("=" * 60)
    print()

    # 步骤1：解析输入文件
    print("步骤 1/4: 解析输入文件...")
    print(f"  输入文件: {input_file}")
    parser = E1Parser(image_dir=str(base_dir / 'input' / 'images'))
    content = parser.parse_file(str(input_file))
    print(f"  ✓ 解析完成")
    print(f"    - 章节数: {len(content.get('chapters', []))}")
    print(f"    - 参考文献数: {len(content.get('references', []))}")
    print()

    # 步骤2：加载样式配置
    print("步骤 2/4: 加载样式配置...")
    print(f"  配置文件: {config_file}")
    style_manager = E1StyleManager(str(config_file))
    print(f"  ✓ 配置加载完成")
    print()

    # 步骤3：格式化文档
    print("步骤 3/4: 格式化文档...")
    formatter = E1Formatter(style_manager)
    doc = formatter.format_document(content)
    print(f"  ✓ 格式化完成")
    print()

    # 步骤4：保存文档
    print("步骤 4/4: 保存文档...")
    print(f"  输出文件: {output_file}")
    doc.save(str(output_file))
    print(f"  ✓ 文档保存成功")
    print()

    # 统计信息
    print("=" * 60)
    print("生成统计:")
    print("=" * 60)
    print(f"摘要: {'✓' if content.get('abstract', {}).get('content') else '×'}")
    print(f"英文摘要: {'✓' if content.get('abstract_en', {}).get('content') else '×'}")
    print(f"目录: ✓")
    print(f"章节: {len(content.get('chapters', []))} 章")

    total_items = 0
    for chapter in content.get('chapters', []):
        total_items += len(chapter.get('content', []))
    print(f"正文内容项: {total_items} 项")

    print(f"参考文献: {len(content.get('references', []))} 条")
    print(f"致谢: {'✓' if content.get('acknowledgements') else '×'}")
    print(f"附录: {'✓' if content.get('appendix') else '×'}")
    print()

    # 自检清单
    print("=" * 60)
    print("实现功能自检清单:")
    print("=" * 60)
    impl_features = style_manager.get_config().get('implementation_features', {})
    print(f"目录可点击链接: {'✓' if impl_features.get('toc_clickable_links') else '×'}")
    print(f"目录自动页码: {'✓' if impl_features.get('toc_auto_page_numbers') else '×'}")
    print(f"目录链接黑色: {'✓' if impl_features.get('toc_link_color') else '×'}")
    print(f"图表SEQ字段: {'✓' if impl_features.get('figure_table_seq_fields') else '×'}")
    print(f"参考文献引用链接: {'✓' if impl_features.get('reference_citation_links') else '×'}")
    print(f"公式OMML格式: {'✓' if impl_features.get('formula_omml_format') else '×'}")
    print(f"中英文字体分离: {'✓' if impl_features.get('font_separate_cjk_latin') else '×'}")
    print(f"页眉独立设置: {'✓' if impl_features.get('header_unlink_from_previous') else '×'}")
    print(f"图片缺失占位符: {'✓' if impl_features.get('missing_image_placeholder') else '×'}")
    print()

    print("=" * 60)
    print("✓ 文档生成成功！")
    print("=" * 60)
    print()

    return 0


if __name__ == '__main__':
    sys.exit(main())
