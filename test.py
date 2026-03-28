#!/usr/bin/env python3
"""将 .docx 转为 Markdown：mammoth（docx→HTML，图片落盘）→ markdownify（HTML→MD）。"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import mammoth
from markdownify import markdownify as html_to_md


def main() -> None:
    parser = argparse.ArgumentParser(description="DOCX 转 Markdown")
    parser.add_argument(
        "docx",
        nargs="?",
        default="新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版).docx",
        help="输入的 .docx 路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="输出的 .md 路径（默认同目录、与 docx 同名）",
    )
    args = parser.parse_args()

    docx_path = Path(args.docx).resolve()
    if not docx_path.is_file():
        sys.exit(f"文件不存在: {docx_path}")

    out_md = (
        Path(args.output).resolve()
        if args.output
        else docx_path.with_suffix(".md")
    )
    media_dir = out_md.parent / f"{out_md.stem}_media"
    media_dir.mkdir(parents=True, exist_ok=True)

    class ImageWriter:
        def __init__(self, directory: Path, rel_prefix: str) -> None:
            self._directory = directory
            self._rel_prefix = rel_prefix
            self._n = 1

        def __call__(self, element):  # mammoth.images.Image
            ext = element.content_type.partition("/")[2] or "bin"
            filename = f"{self._n}.{ext}"
            dest = self._directory / filename
            with dest.open("wb") as out_f:
                with element.open() as in_f:
                    shutil.copyfileobj(in_f, out_f)
            self._n += 1
            return {"src": f"{self._rel_prefix}{filename}"}

    rel_prefix = f"{media_dir.name}/"
    convert_image = mammoth.images.img_element(
        ImageWriter(media_dir, rel_prefix)
    )

    with docx_path.open("rb") as f:
        result = mammoth.convert_to_html(f, convert_image=convert_image)

    for msg in result.messages:
        print(msg, file=sys.stderr)

    md = html_to_md(
        result.value,
        heading_style="ATX",
        strip=["script", "style"],
    )
    out_md.write_text(md, encoding="utf-8")
    print(f"已写入: {out_md}")
    print(f"图片目录: {media_dir}")


if __name__ == "__main__":
    main()
