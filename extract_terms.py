#!/usr/bin/env python3
"""从国家标准 PDF 中提取“术语和定义”章节。

每条术语输出为:
    {"id": "2.1", "term_zh": "床铺面", "term_en": "bed base", "term_def": "放置床垫或床褥的支撑结构。"}

用法:
    python extract_terms.py input.pdf [-o output.json] [--csv output.csv]
"""

import argparse
import csv
import json
import re
import sys

import pdfplumber
from pdfminer.fontmetrics import FONT_METRICS

# 部分标准 PDF 的西文字体未嵌入字宽,单词之间也没有空格字符,
# 因此借用 Times-Bold 的标准字宽估算字符间距,间距足够大时补回空格。
_TIMES_WIDTHS = FONT_METRICS["Times-Bold"][1]
_SPACE_GAP_RATIO = 0.28  # 间距超过字号的 28% 视为单词间空格(词内间距一般 <0.22,词间 >0.35)

_CJK = r"㐀-鿿豈-﫿"
_HEADER_RE = re.compile(r"^GB(/[TZ])?\s*[\d.]+[—-]\d{4}$")
_PAGE_NO_RE = re.compile(r"^[\dⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+$")


def _is_ascii_alnum(ch):
    return ch.isascii() and ch.isalnum()


def _line_text(chars):
    """把一行字符拼成文本,并在西文单词之间补回丢失的空格。"""
    chars = sorted((c for c in chars if not c["text"].isspace()), key=lambda c: c["x0"])
    out = []
    prev = None
    for c in chars:
        if prev is not None:
            width = _TIMES_WIDTHS.get(prev["text"], 500) * prev["size"] / 1000
            gap = c["x0"] - prev["x0"] - width
            if (_is_ascii_alnum(prev["text"]) or _is_ascii_alnum(c["text"])) and gap > _SPACE_GAP_RATIO * c["size"]:
                # 中文与西文之间也补空格,方便后面拆分中英文术语
                out.append(" ")
        out.append(c["text"])
        prev = c
    return "".join(out).strip()


def _page_lines(page):
    """按纵坐标把字符聚成行。

    某些字体(如标点)的基线有约 0.7 倍字号的偏移,所以容差取 0.8 倍字号,
    行距一般在 1.4 倍字号以上,不会把相邻两行合并。
    """
    chars = sorted(page.chars, key=lambda c: (c["top"], c["x0"]))
    lines = []
    for c in chars:
        if lines and c["top"] - lines[-1][0] <= 0.8 * c["size"]:
            lines[-1][1].append(c)
        else:
            lines.append([c["top"], [c]])
    texts = []
    for _, line_chars in lines:
        text = _line_text(line_chars)
        if text and not _HEADER_RE.match(text) and not _PAGE_NO_RE.match(text):
            texts.append(text)
    return texts


def _join(parts):
    """拼接跨行文本:中文直接相连,西文之间保留空格。"""
    text = ""
    for part in parts:
        if text and re.search(r"[A-Za-z0-9,.;]$", text) and re.match(r"[A-Za-z0-9(]", part):
            text += " "
        text += part
    # 去掉中文与中文、中文与中文标点之间多余的空格
    text = re.sub(rf"(?<=[{_CJK}　-〿＀-￯])\s+(?=[{_CJK}　-〿＀-￯])", "", text)
    return text


def _split_term(line):
    """把 “床铺面 bed base” 拆成中文术语和英文术语。"""
    m = re.match(rf"^(.*[{_CJK}）)])\s+([A-Za-z].*)$", line)
    if m:
        return m.group(1).replace(" ", ""), m.group(2).strip()
    return line.replace(" ", ""), ""


def extract_terms(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        lines = [line for page in pdf.pages for line in _page_lines(page)]

    # 定位“N 术语和定义”章节(跳过目录、前言中的引用)
    chapter = None
    start = None
    for i, line in enumerate(lines):
        m = re.match(r"^(\d+)\s*术语和定义$", line.replace(" ", ""))
        if m:
            chapter, start = m.group(1), i + 1
    if chapter is None:
        raise ValueError("未找到“术语和定义”章节")

    next_chapter_re = re.compile(rf"^{int(chapter) + 1}\s*[{_CJK}]")
    id_re = re.compile(rf"^{chapter}(\.\d+)+$")

    terms = []
    current = None
    for line in lines[start:]:
        compact = line.replace(" ", "")
        if next_chapter_re.match(line):
            break
        if id_re.match(compact):
            current = {"id": compact, "term_zh": "", "term_en": "", "def_lines": []}
            terms.append(current)
        elif current is None:
            continue  # 引导语,如“下列术语和定义适用于本文件。”
        elif not current["term_zh"]:
            current["term_zh"], current["term_en"] = _split_term(line)
        else:
            current["def_lines"].append(line)

    return [
        {
            "id": t["id"],
            "term_zh": t["term_zh"],
            "term_en": t["term_en"],
            "term_def": _join(t["def_lines"]),
        }
        for t in terms
    ]


def main(argv=None):
    parser = argparse.ArgumentParser(description="提取标准 PDF 中的“术语和定义”")
    parser.add_argument("pdf", help="输入 PDF 文件")
    parser.add_argument("-o", "--output", help="输出 JSON 文件(默认打印到标准输出)")
    parser.add_argument("--csv", help="同时输出 CSV 文件")
    args = parser.parse_args(argv)

    terms = extract_terms(args.pdf)
    data = json.dumps(terms, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(data + "\n")
    else:
        print(data)

    if args.csv:
        with open(args.csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["id", "term_zh", "term_en", "term_def"])
            writer.writeheader()
            writer.writerows(terms)


if __name__ == "__main__":
    sys.exit(main())
