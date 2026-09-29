"""正規化: 見出し階層・表・箇条書き・空白を一定ルールで整える（ストリーム処理）.

ルール
- 見出しは 1 段ずつしか深くならない（H1 → H3 は H1 → H2 に詰める）
- 空の段落・空の見出しは捨てる。連続する空白は 1 つにまとめ、行頭行末を除去
- 表は矩形にそろえ（不足セルは空文字）、全列が空の列は削除する
- 箇条書きのレベルは直前の項目 +1 までに制限する
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Iterator

from .model import Block, FootnoteDef, Heading, ListItem, Paragraph, Quote, Table

_WS = re.compile(r"[ \t　 ]+")


def clean_text(text: str) -> str:
    lines = [_WS.sub(" ", line).strip() for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    return "\n".join(line for line in lines if line)


def _normalize_table(t: Table) -> Table | None:
    rows = [[clean_text(c) for c in r] for r in t.rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        return None
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    keep = [i for i in range(width) if any(r[i] for r in rows)]
    rows = [[r[i] for i in keep] for r in rows]
    return Table(rows=rows, header=t.header)


def normalize(blocks: Iterable[Block]) -> Iterator[Block]:
    prev_heading = 0
    prev_list_level = -1
    for b in blocks:
        if isinstance(b, Heading):
            text = clean_text(b.text).replace("\n", " ")
            if not text:
                continue
            level = max(1, min(b.level, prev_heading + 1, 6))
            prev_heading = level
            prev_list_level = -1
            yield Heading(level, text, b.synthetic)
        elif isinstance(b, Paragraph):
            text = clean_text(b.text)
            if text:
                yield Paragraph(text, b.synthetic)
            prev_list_level = -1
        elif isinstance(b, ListItem):
            text = clean_text(b.text).replace("\n", " ")
            if not text:
                continue
            level = max(0, min(b.level, prev_list_level + 1))
            prev_list_level = level
            yield ListItem(text, level, b.ordered)
        elif isinstance(b, Table):
            t = _normalize_table(b)
            if t:
                yield t
            prev_list_level = -1
        elif isinstance(b, Quote):
            text = clean_text(b.text)
            if text:
                yield Quote(text, b.label)
            prev_list_level = -1
        elif isinstance(b, FootnoteDef):
            text = clean_text(b.text)
            yield FootnoteDef(b.label, text or "（空）")
        else:
            yield b
