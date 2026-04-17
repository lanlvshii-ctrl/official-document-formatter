#!/usr/bin/env python3
"""
extract_document.py
将源文档（.docx / .doc / .pdf / .md / .txt / 图片）提取为 raw.md
"""

import argparse
import platform
import re
import subprocess
import sys
import tempfile
from pathlib import Path

# 可选依赖
try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from PIL import Image
    import pytesseract
except ImportError:
    Image = None
    pytesseract = None


def _table_to_markdown(table) -> str:
    """将 python-docx 表格转换为 Markdown 表格字符串"""
    rows = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    max_cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < max_cols:
            r.append("")
    lines = []
    lines.append("| " + " | ".join(rows[0]) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def extract_from_docx(path: Path) -> str:
    if Document is None:
        raise RuntimeError("缺少 python-docx，请执行: pip install python-docx")
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn as docx_qn

    doc = Document(str(path))
    items = []
    for child in doc.element.body.iterchildren():
        if child.tag == docx_qn("w:p"):
            p = Paragraph(child, doc)
            if p.text and p.text.strip():
                items.append(p.text.strip())
        elif child.tag == docx_qn("w:tbl"):
            tbl = Table(child, doc)
            md = _table_to_markdown(tbl)
            if md:
                items.append(md)
    return "\n\n".join(items)


def extract_from_doc(path: Path) -> str:
    """将 .doc 转为临时 .docx 再提取。macOS 使用 textutil，其他平台使用 pandoc。"""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        tmp_docx = tmpdir / "temp.docx"

        if platform.system() == "Darwin":
            # macOS: textutil 是内置工具
            result = subprocess.run(
                ["textutil", "-convert", "docx", str(path), "-output", str(tmp_docx)],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                raise RuntimeError(f"textutil 转换 .doc 失败: {result.stderr}")
        else:
            # Windows/Linux: 使用 pandoc
            try:
                result = subprocess.run(
                    ["pandoc", "-f", "doc", "-t", "docx", "-o", str(tmp_docx), str(path)],
                    capture_output=True, text=True
                )
                if result.returncode != 0:
                    raise RuntimeError(f"pandoc 转换 .doc 失败: {result.stderr}")
            except FileNotFoundError:
                raise RuntimeError(
                    "处理 .doc 文件需要 pandoc。请安装：\n"
                    "  - Windows: https://pandoc.org/installing.html\n"
                    "  - Linux: sudo apt install pandoc\n"
                    "或者将 .doc 另存为 .docx 后再使用本工具。"
                )
        return extract_from_docx(tmp_docx)


def is_new_paragraph_marker(line: str) -> bool:
    """判断一行是否是新段落的开头"""
    markers = [
        r"^[一二三四五六七八九十百千万]+、",
        r"^（[一二三四五六七八九十百千万]+）",
        r"^\d+[．\.]",
        r"^（\d+）",
        r"^第[一二三四五六七八九十百千万]+章",
        r"^附件[：:]",
        r"^关于",
    ]
    for pat in markers:
        if re.match(pat, line):
            return True
    return False


def clean_pdf_breaks(text: str) -> str:
    """修正 PDF 提取中常见的过度分段"""
    lines = text.splitlines()
    result = []
    for line in lines:
        line = line.strip()
        if not line:
            if result and result[-1] != "":
                result.append("")
            continue
        # 如果上一行不以句末标点结束，且当前行不是新段落开头，则合并
        if result and result[-1] and not result[-1].endswith((
            "。", "；", "：", "！", "？", "…", ".", ";", ":", "!", "?"
        )):
            if not is_new_paragraph_marker(line):
                result[-1] += line
                continue
        result.append(line)
    # 合并连续空行
    text = "\n".join(result)
    while "\n\n\n" in text:
        text = text.replace("\n\n\n", "\n\n")
    return text


def _extract_pdf_tables(page) -> list[str]:
    """尝试提取 PDF 页面中的表格为 Markdown 表格（需 PyMuPDF >= 1.23）"""
    tables = []
    try:
        tabs = page.find_tables()
        for tab in tabs:
            rows = tab.extract()
            if not rows:
                continue
            max_cols = max(len(r) for r in rows)
            lines = []
            lines.append("| " + " | ".join(str(c or "") for c in rows[0]) + " |")
            lines.append("| " + " | ".join(["---"] * max_cols) + " |")
            for r in rows[1:]:
                lines.append("| " + " | ".join(str(c or "") for c in r) + " |")
            tables.append("\n".join(lines))
    except Exception:
        pass
    return tables


def extract_from_pdf(path: Path) -> str:
    if fitz is None:
        raise RuntimeError("缺少 PyMuPDF，请执行: pip install pymupdf")
    doc = fitz.open(str(path))
    paragraphs = []
    all_tables = []
    for page in doc:
        text = page.get_text().strip()
        if text:
            paragraphs.append(text)
        tables = _extract_pdf_tables(page)
        all_tables.extend(tables)
    doc.close()
    full_text = "\n\n".join(paragraphs)
    result = clean_pdf_breaks(full_text)
    if all_tables:
        result = result + "\n\n" + "\n\n".join(all_tables)
    return result


def read_text_with_fallback(path: Path) -> str:
    """尝试多种编码读取文本文件"""
    encodings = ["utf-8", "gbk", "gb2312", "gb18030", "big5", "latin-1"]
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    # 最后用 utf-8 忽略错误
    return path.read_text(encoding="utf-8", errors="replace")


def extract_from_txt(path: Path) -> str:
    return read_text_with_fallback(path)


def extract_from_md(path: Path) -> str:
    return read_text_with_fallback(path)


def extract_from_image(path: Path) -> str:
    if Image is None or pytesseract is None:
        raise RuntimeError(
            "缺少 OCR 依赖，请执行: pip install pytesseract pillow\n"
            "并安装 Tesseract-OCR 引擎及中文语言包:\n"
            "  - macOS: brew install tesseract tesseract-lang\n"
            "  - Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "  - Ubuntu/Debian: sudo apt install tesseract-ocr tesseract-ocr-chi-sim"
        )
    img = Image.open(str(path))
    text = pytesseract.image_to_string(img, lang="chi_sim+eng")
    return text.strip()


def classify_and_extract(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_from_docx(path)
    elif suffix == ".doc":
        return extract_from_doc(path)
    elif suffix == ".pdf":
        return extract_from_pdf(path)
    elif suffix == ".txt":
        return extract_from_txt(path)
    elif suffix in (".md", ".markdown"):
        return extract_from_md(path)
    elif suffix in (".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif"):
        return extract_from_image(path)
    else:
        raise ValueError(f"不支持的文件格式: {suffix}")


def count_chinese_chars(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def main():
    parser = argparse.ArgumentParser(description="将源文档提取为 raw.md")
    parser.add_argument("input", help="源文档路径")
    parser.add_argument("-o", "--output", help="输出 raw.md 路径（默认与源文件同目录）")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        sys.exit(1)

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.parent / f"{input_path.stem}_raw.md"

    try:
        text = classify_and_extract(input_path)
    except Exception as e:
        print(f"❌ 提取失败: {e}")
        sys.exit(1)

    # 统一换行符，去除多余空行
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    while "\n\n\n" in text:
        text = text.replace("\n\n\n", "\n\n")

    output_path.write_text(text, encoding="utf-8")

    cn_chars = count_chinese_chars(text)
    print(f"✅ 提取完成: {output_path}")
    print(f"   中文字符数: {cn_chars}")

    if cn_chars < 50:
        print("⚠️  警告: 中文字符数过少，请检查源文件是否为扫描件/图片，可能需要 OCR 处理。")
    elif cn_chars < 200:
        print("ℹ️  提示: 提取内容较短，建议人工核对。")


if __name__ == "__main__":
    main()
