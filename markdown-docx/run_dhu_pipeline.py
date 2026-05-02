#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
from pathlib import Path

from fill_template import DEFAULT_SCHEMA, DEFAULT_TEMPLATE, fill_document
from optimize_clean_template import optimize_clean_template


SCRIPT_DIR = Path(__file__).resolve().parent


def main() -> None:
    parser = argparse.ArgumentParser(description="一键执行东华论文模板优化与 Markdown 直填")
    parser.add_argument("markdown", help="输入 Markdown 路径")
    parser.add_argument("-o", "--output", help="输出 docx 路径")
    parser.add_argument("-m", "--meta", help="元数据 JSON 路径")
    parser.add_argument("-t", "--template", default=str(DEFAULT_TEMPLATE), help="干净模板路径")
    parser.add_argument("-s", "--schema", default=str(DEFAULT_SCHEMA), help="模板 schema 路径")
    parser.add_argument("--skip-optimize", action="store_true", help="跳过模板优化步骤")
    args = parser.parse_args()

    markdown_path = Path(args.markdown).resolve()
    template_path = Path(args.template).resolve()
    schema_path = Path(args.schema).resolve()
    metadata_path = Path(args.meta).resolve() if args.meta else None
    output_path = Path(args.output).resolve() if args.output else SCRIPT_DIR / "outputs" / f"{markdown_path.stem}_东华直填.docx"

    if not args.skip_optimize:
        optimize_clean_template(template_path, schema_path)
        print(f"已优化干净模板: {template_path}")

    fill_document(template_path, markdown_path, output_path, schema_path, metadata_path)
    print(f"已生成: {output_path}")


if __name__ == "__main__":
    main()
