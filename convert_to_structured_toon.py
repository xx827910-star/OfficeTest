#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 JSON 版格式数据渲染为“结构化 TOON”文本。

目标：
1. 保留 TOON 的缩进/无花括号语法以维持较低 token 数。
2. 对列表类字段（items/indexes/tab_stops 等）使用显式 key-value 形式，
   避免 LLM 需要依赖逗号位置去推断字段顺序，从而降低误读概率。
3. 输出到独立目录，便于与现有 TOON 版本对比。
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Iterable, List

INDENT = "  "


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert JSON format exports into a structured TOON variant."
    )
    parser.add_argument(
        "--input-dir",
        default="json_output",
        help="目录，包含 format_data_vXX.json (默认: json_output)",
    )
    parser.add_argument(
        "--output-dir",
        default="toon_output_structured",
        help="输出目录 (默认: toon_output_structured)",
    )
    parser.add_argument(
        "--pattern",
        default="format_data_v*.json",
        help="匹配的 JSON 文件模式 (默认: format_data_v*.json)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    json_paths = sorted(input_dir.glob(args.pattern))
    if not json_paths:
        raise SystemExit(f"未在 {input_dir} 找到匹配 {args.pattern} 的 JSON 文件")

    for json_path in json_paths:
        data = json.loads(json_path.read_text(encoding="utf-8"))
        toon_text = render_document(data)
        target_path = output_dir / (json_path.stem + ".toon")
        target_path.write_text(toon_text, encoding="utf-8")
        rel_path = target_path.resolve().relative_to(Path.cwd())
        print(f"[OK] {json_path.name} -> {rel_path}")


def render_document(data: Any) -> str:
    """将任意 JSON 数据渲染为结构化 TOON 字符串。"""
    lines = render_object(data, "")
    return "\n".join(lines) + "\n"


def render_object(value: Any, indent: str) -> List[str]:
    """递归渲染字典或值。"""
    if isinstance(value, dict):
        return render_dict(value, indent)
    return [indent + format_scalar(value)]


def render_dict(data: dict, indent: str) -> List[str]:
    lines: List[str] = []
    for key, value in data.items():
        lines.extend(render_key_value(key, value, indent))
    return lines


def render_key_value(key: str, value: Any, indent: str) -> List[str]:
    prefix = f"{indent}{key}:"

    if isinstance(value, dict):
        if not value:
            return [prefix + " {}"]
        lines = [prefix]
        lines.extend(render_dict(value, indent + INDENT))
        return lines

    if isinstance(value, list):
        if not value:
            return [prefix + " []"]
        lines = [prefix]
        lines.extend(render_list(value, indent + INDENT))
        return lines

    return [prefix + " " + format_scalar(value)]


def render_list(items: Iterable[Any], indent: str) -> List[str]:
    lines: List[str] = []
    for item in items:
        if isinstance(item, dict):
            lines.append(f"{indent}-")
            lines.extend(render_dict(item, indent + INDENT))
        elif isinstance(item, list):
            lines.append(f"{indent}-")
            lines.extend(render_list(item, indent + INDENT))
        else:
            lines.append(f"{indent}- {format_scalar(item)}")
    return lines


def format_scalar(value: Any) -> str:
    if value is None:
        return '""'
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, str):
        return format_string(value)
    return json.dumps(value, ensure_ascii=False)


def format_string(text: str) -> str:
    if text == "":
        return '""'
    if any(ch in text for ch in (",", ":", "\n", '"')):
        return json.dumps(text, ensure_ascii=False)
    if text.strip() != text:
        return json.dumps(text, ensure_ascii=False)
    return text


if __name__ == "__main__":
    main()
