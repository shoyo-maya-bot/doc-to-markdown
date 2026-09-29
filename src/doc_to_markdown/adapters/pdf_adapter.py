"""PDF → 中間表現（pdfplumber。スキャンページは Tesseract OCR）.

- ページごとに表（罫線のある表は find_tables）と、表の外側のテキストを分けて抽出
- 見出しの推定:
  - 本文の代表サイズ（文字数の最頻値）の 1.15 倍以上の行 → 文書内で大きい順に H2 / H3 / H4
  - 本文サイズでも「第N章 / 第N節 / 第N条」で始まる短い行、太字だけの短い行 → 見出し
- 段落: 行間が詰まっていて、前の行が右端近くまで埋まっている（折り返し）行を結合
  （日本語は空白なし、英数字は空白 1 つで結合）
- 罫線のない表: 広い語間で 3 列以上に分かれる行が同じ列数で 2 行以上続けば表とみなす
- 2 段組み: ページ中央付近に縦の空白帯（ガター）があれば、左段 → 右段の順に読む
- 繰り返しのヘッダー・フッター: 半数以上のページの上端・下端に同じ行（ページ番号は数字を無視して比較）
  があれば除去する（--keep-running-headers で残す）
- スキャンページ: 文字の無いページは Tesseract で OCR（--ocr auto/always/never、--ocr-lang）
- 出典トレースのため各ページ先頭に <!-- page N --> を入れる
- 大きな PDF でもメモリが増え続けないよう、1 ページずつ処理してキャッシュを解放する
"""

from __future__ import annotations

import bisect
import re
import statistics
import subprocess
from collections import Counter
from collections.abc import Iterator
from pathlib import Path

import pdfplumber
from devtools_common.executables import TESSERACT, find_executable, not_found_message
from devtools_common.tempfiles import secure_tempdir

from ..assets import AssetWriter
from ..model import Block, Comment, Document, Heading, ListItem, Options, Paragraph, SourceStats, Table, count_chars

# \uf0xx は Symbol / Wingdings フォントの記号（Word の箇条書きを PDF 化すると現れる）
_BULLET = re.compile(r"^\s*([・•●○■□◆◇▪◦‣\-–*-])\s*(.+)$")
_BULLET_ONLY = re.compile(r"^[・•●○■□◆◇▪◦‣\-–*-]$")
_ORDERED = re.compile(r"^\s*(\d{1,3})[.)．）]\s+(.+)$")
_ARTICLE = re.compile(r"^第[0-9０-９一二三四五六七八九十百]+\s*([章節条項])")
_ASCII_END = re.compile(r"[A-Za-z0-9,.;:)]$")
_BOLD_FONT = re.compile(r"bold|black|heavy|semibold|demi|-w[6-9]\b|w[6-9]$", re.I)
HEADING_RATIO = 1.15
COLUMN_GAP = 1.5  # 文字サイズのこの倍率を超える語間は列の区切りとみなす（罫線なし表の検出）
MIN_TABLE_COLUMNS = 3
EDGE_ZONE = 0.08  # ページ上端・下端のこの割合の範囲をヘッダー・フッター候補とする
SCANNED_MAX_CHARS = 10


def _join(a: str, b: str) -> str:
    return f"{a} {b}" if _ASCII_END.search(a) and b[:1].isascii() else a + b


def _signature(text: str) -> str:
    return re.sub(r"\d+", "#", re.sub(r"\s+", "", text))


# ---------------------------------------------------------------------------
# 語 → 行
# ---------------------------------------------------------------------------
def _words(page, tables_bbox):
    def outside(obj):
        x0, top, x1, bottom = obj.get("x0", 0), obj.get("top", 0), obj.get("x1", 0), obj.get("bottom", 0)
        return not any(
            bx0 <= x0 and x1 <= bx1 and btop <= top and bottom <= bbottom for bx0, btop, bx1, bbottom in tables_bbox
        )

    return page.filter(outside).extract_words(
        extra_attrs=["size", "fontname"], keep_blank_chars=True, use_text_flow=True
    )


def _lines(words) -> list[dict]:
    """語を行にまとめる: 縦に重なる語を同じ行とし、行内は左から順に並べる."""
    rows: list[dict] = []
    for w in sorted(words, key=lambda w: w["top"]):
        center = (w["top"] + w["bottom"]) / 2
        # 記号だけフォントサイズが違う場合も同じ行にするため、中心が行の上下端の内側なら同じ行
        row = next((r for r in reversed(rows[-4:]) if r["top"] - 1 <= center <= r["bottom"] + 1), None)
        if row is None:
            rows.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
        else:
            row["words"].append(w)
    lines: list[dict] = []
    for r in rows:
        ws = sorted(r["words"], key=lambda w: w["x0"])
        first = ws[0]
        ln = {
            "text": first["text"],
            "top": min(w["top"] for w in ws),
            "bottom": max(w["bottom"] for w in ws),
            "x0": first["x0"],
            "x1": first["x1"],
            "size": first["size"],
            "cells": [first["text"]],
            "bold": bool(_BOLD_FONT.search(first.get("fontname", ""))),
        }
        for w in ws[1:]:
            gap = w["x0"] - ln["x1"]
            sep = " " if gap > w["size"] * 0.25 else ""
            ln["text"] += sep + w["text"]
            if gap > w["size"] * COLUMN_GAP:
                ln["cells"].append(w["text"])
            else:
                ln["cells"][-1] += sep + w["text"]
            ln["x1"] = max(ln["x1"], w["x1"])
            if not _BULLET_ONLY.match(w["text"]):
                starts_with_bullet = _BULLET_ONLY.match(ln["text"].split()[0]) is not None
                ln["size"] = w["size"] if starts_with_bullet and len(ln["cells"]) == 1 else max(ln["size"], w["size"])
                ln["bold"] = ln["bold"] and bool(_BOLD_FONT.search(w.get("fontname", "")))
        if _BULLET_ONLY.match(ln["cells"][0]) and len(ln["cells"]) > 1:
            ln["cells"] = [ln["cells"][0] + " " + ln["cells"][1], *ln["cells"][2:]]
        if _BULLET_ONLY.match(ln["text"].split()[0]) and len(ws) > 1:
            ln["bold"] = all(_BOLD_FONT.search(w.get("fontname", "")) for w in ws[1:])
        lines.append(ln)
    lines.sort(key=lambda ln: (ln["top"], ln["x0"]))
    return lines


def find_gutter(words, width: float) -> float | None:
    """2 段組みの段間（縦の空白帯）の x 座標. 見つからなければ None."""
    if len(words) < 12:
        return None
    best = None
    for i in range(30, 71, 2):
        x = width * i / 100
        crossing = sum(1 for w in words if w["x0"] < x < w["x1"])
        left = [w for w in words if w["x1"] <= x]
        right = [w for w in words if w["x0"] >= x]
        if len(left) < len(words) * 0.15 or len(right) < len(words) * 0.15:
            continue
        if crossing > len(words) * 0.03:
            continue
        # 左右が同じ高さに並んでいること（単なる右寄せの行ではない）
        left_rows = {round(w["top"] / 4) for w in left}
        shared = sum(1 for w in right if round(w["top"] / 4) in left_rows)
        if shared < 3:
            continue
        score = (crossing, abs(i - 50))
        if best is None or score < best[0]:
            best = (score, x)
    return best[1] if best else None


# ---------------------------------------------------------------------------
# ページの組み立て
# ---------------------------------------------------------------------------
class _PageBuilder:
    """1 ページ分の行・表を Block 列に組み立てる（1 段組み単位）."""

    def __init__(self, lines: list[dict], level_of: dict[float, int], body_size: float, st: SourceStats):
        self.lines, self.level_of, self.body_size, self.st = lines, level_of, body_size, st
        gaps = [b["top"] - a["bottom"] for a, b in zip(lines, lines[1:], strict=False) if b["top"] > a["bottom"]]
        self.gap_limit = (statistics.median(gaps) * 1.8) if gaps else body_size
        # 折り返し判定の右端は本文サイズの行だけで決める（大きな見出しに引っ張られないように）
        body = [ln["x1"] for ln in lines if abs(ln["size"] - body_size) < 1.0]
        self.right_edge = max(body or [ln["x1"] for ln in lines] or [0])
        # 箇条書きの字下げ位置（左端 x）の段階をレベルにする
        self.indents = sorted({round(ln["x0"]) for ln in lines if _BULLET.match(ln["text"].strip())})
        self.items: list[tuple[float, Block]] = []
        self.para: list[str] = []
        self.para_top = 0.0
        self.run: list[dict] = []  # 罫線なし表の候補行
        self.min_level = min(level_of.values(), default=2)
        self.max_level = max(level_of.values(), default=1)

    def flush_para(self) -> None:
        if self.para:
            text = self.para[0]
            for nxt in self.para[1:]:
                text = _join(text, nxt)
            self.st.add_text(text)
            self.items.append((self.para_top, Paragraph(text)))
            self.para = []

    def flush_run(self) -> None:
        if len(self.run) >= 2:
            rows = [[c.strip() for c in r["cells"]] for r in self.run]
            self.st.tables += 1
            self.st.table_rows += len(rows)
            self.st.add_text("".join("".join(r) for r in rows))
            self.items.append((self.run[0]["top"], Table(rows=rows, header=True)))
        else:
            for r in self.run:
                self.st.add_text(r["text"])
                self.items.append((r["top"], Paragraph(r["text"].strip())))
        self.run = []

    def heading_level(self, ln: dict, text: str) -> int | None:
        level = self.level_of.get(round(ln["size"], 1))
        if level and len(text) <= 80:
            return level
        if len(text) > 40 or text.endswith(("。", "、", ".", ",")):
            return None
        m = _ARTICLE.match(text)
        if m:
            return {"章": 2, "節": 3}.get(m.group(1), min(self.max_level + 1, 4))
        if ln.get("bold") and abs(ln["size"] - self.body_size) < 0.6:
            return min(self.max_level + 1, 4)
        return None

    def build(self, tables) -> list[tuple[float, Block]]:
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
            bullet, ordered = _BULLET.match(text), _ORDERED.match(text)
            level = self.heading_level(ln, text)
            # 記号付きの行は文字が大きくても箇条書き（スライドの本文など）。番号付きで大きい行は見出し
            if bullet or (ordered and not level):
                self.flush_para()
                self.st.list_items += 1
                m = bullet or ordered
                self.st.add_text(m.group(2))
                indent = bisect.bisect_left(self.indents, round(ln["x0"]) - 3)
                self.items.append((ln["top"], ListItem(m.group(2), indent, bool(ordered))))
            elif level:
                self.flush_para()
                self.st.headings += 1
                self.st.add_text(text)
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
            self.st.add_text("".join("".join(r) for r in rows))
            self.items.append((t.bbox[1], Table(rows=rows, header=True)))
        self.items.sort(key=lambda it: it[0])
        return self.items


# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------
class _Ocr:
    def __init__(self, opts: Options, doc: Document):
        self.opts, self.doc = opts, doc
        self.exe = find_executable(TESSERACT) if opts.ocr != "never" else None
        self.lang: str | None = None
        self.pages: list[int] = []
        self.skipped: list[int] = []

    def resolve_lang(self) -> str | None:
        if self.lang is not None or not self.exe:
            return self.lang
        r = subprocess.run([self.exe, "--list-langs"], capture_output=True, text=True, timeout=60)
        available = {ln.strip() for ln in r.stdout.splitlines()[1:] if ln.strip()}
        wanted = [x for x in self.opts.ocr_lang.split("+") if x]
        usable = [x for x in wanted if x in available]
        missing = [x for x in wanted if x not in available]
        if missing:
            self.doc.warnings.append(
                f"OCR の言語データ {', '.join(missing)} がインストールされていません"
                f"（使用: {'+'.join(usable) or 'なし'}）。日本語の読み取りには jpn が必要です"
            )
        self.lang = "+".join(usable)
        return self.lang or None

    def read(self, page, number: int) -> str | None:
        if not self.exe:
            self.skipped.append(number)
            return None
        lang = self.resolve_lang()
        if not lang:
            self.skipped.append(number)
            return None
        with secure_tempdir("doc2md-ocr-") as tmp:
            png = tmp / "page.png"
            page.to_image(resolution=self.opts.ocr_dpi).original.save(png)
            r = subprocess.run(
                [self.exe, str(png), "stdout", "-l", lang, "--psm", "3"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                timeout=300,
            )
        if r.returncode != 0:
            self.skipped.append(number)
            return None
        self.pages.append(number)
        return r.stdout

    def finish(self) -> None:
        if self.pages:
            self.doc.warnings.append(
                f"OCR で読み取ったページがあります（{_pages(self.pages)}）。"
                "誤認識の可能性があるため原本と照合してください"
            )
        if self.skipped:
            reason = (
                not_found_message(TESSERACT) if not self.exe and self.opts.ocr != "never" else "OCR を行いませんでした"
            )
            self.doc.warnings.append(f"文字の無いページ（スキャン画像の可能性）: {_pages(self.skipped)}。{reason}")


def _pages(nums: list[int]) -> str:
    return ", ".join(str(n) for n in nums[:20]) + (" …" if len(nums) > 20 else "") + " ページ"


def _ocr_blocks(text: str, st: SourceStats) -> list[Block]:
    blocks: list[Block] = []
    for para in re.split(r"\n\s*\n", text):
        lines = [ln.strip() for ln in para.splitlines() if ln.strip()]
        if not lines:
            continue
        # OCR の行は推測で結合せず、改行を保ったまま出す（誤認識を見つけやすくするため）
        # OCR は日本語の文字間に空白を入れがちなので詰める
        joined = "\n".join(re.sub(r"(?<=[^\x00-\x7f]) (?=[^\x00-\x7f])", "", ln) for ln in lines)
        m = _BULLET.match(joined)
        st.add_text(joined)
        if m:
            st.list_items += 1
            blocks.append(ListItem(m.group(2)))
        else:
            blocks.append(Paragraph(joined))
    return blocks


# ---------------------------------------------------------------------------
def _scan(pdf, opts: Options) -> tuple[Counter, Counter, int, int]:
    """1 回目の走査: 文字サイズの分布・端の行の出現回数・全文字数・ページ数."""
    sizes: Counter[float] = Counter()
    edges: Counter[str] = Counter()
    raw = 0
    for page in pdf.pages:
        chars = page.chars
        raw += count_chars("".join(ch["text"] for ch in chars))
        for ch in chars:
            if ch["text"].strip():
                sizes[round(ch["size"], 1)] += 1
        if opts.remove_running_headers and len(pdf.pages) >= 2 and chars:
            h = page.height
            words = page.extract_words(keep_blank_chars=True)
            for ln in _lines([dict(w, size=w["bottom"] - w["top"], fontname="") for w in words]):
                if ln["top"] < h * EDGE_ZONE or ln["bottom"] > h * (1 - EDGE_ZONE):
                    edges[_signature(ln["text"])] += 1
        page.close()
    return sizes, edges, raw, len(pdf.pages)


def _blocks(path: Path, doc: Document, opts: Options) -> Iterator[Block]:
    st = doc.stats
    yield Heading(1, path.stem, synthetic=True)
    ocr = _Ocr(opts, doc)
    removed_chars = 0
    two_column_pages: list[int] = []
    with pdfplumber.open(str(path)) as pdf:
        sizes, edges, raw, n_pages = _scan(pdf, opts)
        body_size = sizes.most_common(1)[0][0] if sizes else 10.0
        heading_sizes = sorted({s for s in sizes if s >= body_size * HEADING_RATIO}, reverse=True)
        level_of = {s: min(2 + i, 4) for i, s in enumerate(heading_sizes)}
        running = {sig for sig, n in edges.items() if n >= max(2, n_pages * 0.5) and sig}

        for page in pdf.pages:
            number = page.page_number
            yield Comment(f"page {number}")
            text_chars = sum(1 for ch in page.chars if ch["text"].strip())
            scanned = text_chars <= SCANNED_MAX_CHARS and bool(page.images)
            if opts.ocr == "never" and scanned:
                ocr.skipped.append(number)
            if opts.ocr == "always" or (opts.ocr == "auto" and scanned):
                text = ocr.read(page, number)
                if text is not None:
                    yield from _ocr_blocks(text, st)
                    page.close()
                    continue
            tables = page.find_tables()
            words = _words(page, [t.bbox for t in tables])
            h = page.height
            if running:
                removed = [
                    ln
                    for ln in _lines(words)
                    if (ln["top"] < h * EDGE_ZONE or ln["bottom"] > h * (1 - EDGE_ZONE))
                    and _signature(ln["text"]) in running
                ]
                removed_chars += sum(count_chars(ln["text"]) for ln in removed)
                words = [
                    w
                    for w in words
                    if not any(
                        ln["top"] - 1 <= (w["top"] + w["bottom"]) / 2 <= ln["bottom"] + 1
                        and ln["x0"] - 1 <= w["x0"] <= ln["x1"]
                        for ln in removed
                    )
                ]
            gutter = find_gutter(words, page.width)
            items: list[tuple[tuple, Block]] = []
            if gutter is None:
                for top, b in _PageBuilder(_lines(words), level_of, body_size, st).build(tables):
                    items.append(((top,), b))
            else:
                two_column_pages.append(number)
                left = [w for w in words if w["x1"] <= gutter]
                right = [w for w in words if w["x0"] >= gutter]
                span = [w for w in words if w["x0"] < gutter < w["x1"]]
                l_tables = [t for t in tables if t.bbox[2] <= gutter]
                r_tables = [t for t in tables if t.bbox[0] >= gutter]
                s_tables = [t for t in tables if t not in l_tables and t not in r_tables]
                span_items = _PageBuilder(_lines(span), level_of, body_size, st).build(s_tables)
                bounds = sorted(top for top, _ in span_items)
                for col, (ws, ts) in enumerate(((left, l_tables), (right, r_tables))):
                    for top, b in _PageBuilder(_lines(ws), level_of, body_size, st).build(ts):
                        section = bisect.bisect_right(bounds, top)
                        items.append(((section, 1, col, top), b))
                for top, b in span_items:
                    items.append(((bisect.bisect_left(bounds, top), 0, 0, top), b))
            items.sort(key=lambda it: it[0])
            yield from (b for _, b in items)
            page.close()  # ページ単位でキャッシュを解放（大きな PDF 対策）
    ocr.finish()
    doc.raw_text_chars = max(0, raw - removed_chars)
    if running:
        doc.notes.append(f"全ページに繰り返し出る行（ヘッダー・フッター・ページ番号）{len(running)} 種類を除去")
    if two_column_pages:
        doc.notes.append(f"2 段組みとして左段 → 右段の順に読んだページ: {_pages(two_column_pages)}")
    doc.notes.append("PDF の見出し・箇条書きは文字サイズ・太字・行頭記号からの推定")


def convert(path: Path, assets: AssetWriter, opts: Options | None = None) -> Document:
    opts = opts or Options()
    doc = Document()
    doc.blocks = _blocks(path, doc, opts)
    return doc
