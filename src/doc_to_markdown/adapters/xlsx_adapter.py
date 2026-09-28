"""Excel (.xlsx) → 中間表現（openpyxl, read_only でストリーム読み込み）.

- シートごとに見出し（H2）＋表。先頭の非空行をヘッダとみなす
- 数式はキャッシュされた計算結果（data_only=True）を出力する
- 完全に空の行・列は除去する
- 非表示シートは見出しに（非表示）と付けて出力する
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl

from ..assets import AssetWriter
from ..model import Document, Heading, Paragraph, Table


def format_value(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ", timespec="minutes")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    return str(v)


def convert(path: Path, assets: AssetWriter) -> Document:
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    doc = Document()
    st = doc.stats
    doc.blocks.append(Heading(1, path.stem, synthetic=True))
    try:
        for ws in wb.worksheets:
            hidden = getattr(ws, "sheet_state", "visible") != "visible"
            doc.blocks.append(Heading(2, ws.title + ("（非表示）" if hidden else "")))
            st.headings += 1
            rows = []
            for values in ws.iter_rows(values_only=True):
                cells = [format_value(v) for v in values]
                if any(c.strip() for c in cells):
                    rows.append(cells)
            if not rows:
                doc.blocks.append(Paragraph("（空のシート）", synthetic=True))
                continue
            st.tables += 1
            st.table_rows += len(rows)
            for r in rows:
                for c in r:
                    st.add_text(c)
            doc.blocks.append(Table(rows=rows, header=True))
    finally:
        wb.close()
    doc.notes.append("数式はキャッシュ値を出力（未計算のブックは空になる場合あり）。グラフ・図形は対象外")
    return doc
