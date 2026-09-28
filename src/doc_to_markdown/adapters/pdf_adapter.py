"""PDF → 中間表現（pdfplumber）.

- ページごとに表（find_tables）と、表の外側のテキストを分けて抽出
- 文字サイズから見出しを推定: 本文の代表サイズ（文字数の最頻値）の 1.15 倍以上の行を見出しとし、
  文書内で大きい順に H2 / H3 / H4 を割り当てる
- 行を段落に結合する条件: 行間が詰まっている かつ 前の行が右端近くまで埋まっている（折り返し）
  （日本語は空白なし、英数字は空白 1 つで結合）
- 罫線のない表: 広い語間で 3 列以上に分かれる行が同じ列数で 2 行以上続けば表とみなす
- 出典トレースのため各ページ先頭に <!-- page N --> を入れる
"""

from __future__ import annotations

import bisect
import re
import statistics
from collections import Counter
from pathlib import Path

import pdfplumber

from ..assets import AssetWriter
from ..model import Block, Comment, Document, Heading, ListItem, Paragraph, SourceStats, Table

# \uf0xx は Symbol / Wingdings フォントの記号（Word の箇条書きを PDF 化すると現れる）
_BULLET = re.compile(r"^\s*([・•●○■□◆◇▪◦‣\-–*\uf000-\uf0ff])\s*(.+)$")
_BULLET_ONLY = re.compile(r"^[・•●○■□◆◇▪◦‣\-–*\uf000-\uf0ff]$")
HEADING_RATIO = 1.15
COLUMN_GAP = 1.5  # 文字サイズのこの倍率を超える語間は列の区切りとみなす（罫線なし表の検出）
MIN_TABLE_COLUMNS = 3
_ORDERED = re.compile(r"^\s*(\d{1,3})[.)．）]\s+(.+)$")
_ASCII_END = re.compile(r"[A-Za-z0-9,.;:)]$")


def _join(a: str, b: str) -> str:
    return f"{a} {b}" if _ASCII_END.search(a) and b[:1].isascii() else a + b


def _lines(page, tables_bbox):
    def outside(obj):
        x0, top, x1, bottom = obj.get("x0", 0), obj.get("top", 0), obj.get("x1", 0), obj.get("bottom", 0)
        return not any(
            bx0 <= x0 and x1 <= bx1 and btop <= top and bottom <= bbottom for bx0, btop, bx1, bbottom in tables_bbox
        )

    words = page.filter(outside).extract_words(extra_attrs=["size"], keep_blank_chars=True, use_text_flow=True)
    lines: list[dict] = []
    for w in words:
        center = (w["top"] + w["bottom"]) / 2
        # 同じ行: 語の縦中心が行の上下端の内側（記号だけフォントサイズが違う場合も同じ行にする）
        if lines and lines[-1]["top"] - 1 <= center <= lines[-1]["bottom"] + 1:
            ln = lines[-1]
            gap = w["x0"] - ln["x1"]
            ln["text"] += (" " if gap > w["size"] * 0.25 else "") + w["text"]
            if gap > w["size"] * COLUMN_GAP:
                ln["cells"].append(w["text"])
            else:
                ln["cells"][-1] += (" " if gap > w["size"] * 0.25 else "") + w["text"]
            ln["x1"], ln["bottom"] = w["x1"], max(ln["bottom"], w["bottom"])
            if _BULLET_ONLY.match(ln["cells"][0]) and len(ln["cells"]) > 1:
                ln["cells"] = [ln["cells"][0] + " " + ln["cells"][1], *ln["cells"][2:]]
            if not _BULLET_ONLY.match(w["text"]):
                ln["size"] = w["size"] if _BULLET_ONLY.match(ln["text"].split()[0]) else max(ln["size"], w["size"])
        else:
            lines.append(
                {
                    "text": w["text"],
                    "top": w["top"],
                    "bottom": w["bottom"],
                    "x0": w["x0"],
                    "x1": w["x1"],
                    "size": w["size"],
                    "cells": [w["text"]],
                }
            )
    return lines


class _PageBuilder:
    """1 ページ分の行・表を、上から順の Block 列に組み立てる."""

    def __init__(self, lines: list[dict], level_of: dict[float, int], body_size: float, st: SourceStats):
        self.lines, self.level_of, self.st = lines, level_of, st
        gaps = [b["top"] - a["bottom"] for a, b in zip(lines, lines[1:], strict=False) if b["top"] > a["bottom"]]
        self.gap_limit = (statistics.median(gaps) * 1.8) if gaps else body_size
        self.right_edge = max((ln["x1"] for ln in lines), default=0)
        # 箇条書きの字下げ位置（左端 x）の段階をレベルにする
        self.indents = sorted({round(ln["x0"]) for ln in lines if _BULLET.match(ln["text"].strip())})
        self.items: list[tuple[float, Block]] = []
        self.para: list[str] = []
        self.para_top = 0.0
        self.run: list[dict] = []  # 罫線なし表の候補行

    def flush_para(self) -> None:
        if self.para:
            text = self.para[0]
            for nxt in self.para[1:]:
                text = _join(text, nxt)
            self.items.append((self.para_top, Paragraph(text)))
            self.para = []

    def flush_run(self) -> None:
        if len(self.run) >= 2:
            rows = [[c.strip() for c in r["cells"]] for r in self.run]
            self.st.tables += 1
            self.st.table_rows += len(rows)
            self.items.append((self.run[0]["top"], Table(rows=rows, header=True)))
        else:
            self.items.extend((r["top"], Paragraph(r["text"].strip())) for r in self.run)
        self.run = []

    def build(self, tables) -> list[Block]:
        prev_bottom: float | None = None
        prev_x1 = 0.0
        for ln in self.lines:
            text = ln["text"].strip()
            if not text:
                continue
            if len(ln["cells"]) >= MIN_TABLE_COLUMNS and not _BULLET.match(text):
                if self.run and len(self.run[-1]["cells"]) != len(ln["cells"]):
                    self.flush_run()
                self.flush_para()
                self.run.append(ln)
                prev_bottom, prev_x1 = ln["bottom"], 0.0
                continue
            self.flush_run()
            level = self.level_of.get(round(ln["size"], 1))
            bullet, ordered = _BULLET.match(text), _ORDERED.match(text)
            # 記号付きの行は文字が大きくても箇条書き（スライドの本文など）。番号付きで大きい行は見出し
            if bullet or (ordered and not level):
                self.flush_para()
                self.st.list_items += 1
                m = bullet or ordered
                indent = bisect.bisect_left(self.indents, round(ln["x0"]) - 3)
                self.items.append((ln["top"], ListItem(m.group(2), indent, bool(ordered))))
            elif level and len(text) <= 80:
                self.flush_para()
                self.st.headings += 1
                self.items.append((ln["top"], Heading(level, text)))
            else:
                wrapped = prev_x1 >= self.right_edge * 0.85
                if self.para and prev_bottom is not None and (ln["top"] - prev_bottom > self.gap_limit or not wrapped):
                    self.flush_para()
                if not self.para:
                    self.para_top = ln["top"]
                self.para.append(text)
            prev_bottom, prev_x1 = ln["bottom"], ln["x1"]
        self.flush_run()
        self.flush_para()

        for t in tables:
            rows = [[c or "" for c in r] for r in t.extract()]
            if not rows:
                continue
            self.st.tables += 1
            self.st.table_rows += sum(1 for r in rows if any(c.strip() for c in r))
            self.items.append((t.bbox[1], Table(rows=rows, header=True)))
        self.items.sort(key=lambda it: it[0])
        return [b for _, b in self.items]


def convert(path: Path, assets: AssetWriter) -> Document:
    doc = Document()
    st = doc.stats
    doc.blocks.append(Heading(1, path.stem, synthetic=True))
    with pdfplumber.open(str(path)) as pdf:
        pages = []
        sizes: Counter[float] = Counter()
        for page in pdf.pages:
            tables = page.find_tables()
            lines = _lines(page, [t.bbox for t in tables])
            for ln in lines:
                sizes[round(ln["size"], 1)] += len(ln["text"])
            st.add_text("".join(ch["text"] for ch in page.chars))
            pages.append((page.page_number, lines, tables))
        body_size = sizes.most_common(1)[0][0] if sizes else 10.0
        heading_sizes = sorted({s for s in sizes if s >= body_size * HEADING_RATIO}, reverse=True)
        level_of = {s: min(2 + i, 4) for i, s in enumerate(heading_sizes)}
        for number, lines, tables in pages:
            doc.blocks.append(Comment(f"page {number}"))
            doc.blocks.extend(_PageBuilder(lines, level_of, body_size, st).build(tables))
    doc.notes.append("PDF の見出し・箇条書きは文字サイズと行頭記号からの推定")
    return doc
