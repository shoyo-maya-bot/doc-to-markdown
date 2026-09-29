"""中間表現 → Markdown（GFM）. 同じ入力からは常に同じ文字列を返す.

`iter_chunks` はブロックを順に Markdown の断片（段落単位）にして返すので、大きな文書も
ファイルへ逐次書き出せる。断片どうしは空行 1 つで区切る。
"""

from __future__ import annotations

from collections.abc import Iterable, Iterator

from .model import Block, Comment, FootnoteDef, Heading, Image, ListItem, Paragraph, Quote, StreamTable, Table


def _cell(text: str) -> str:
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def render_table(t: Table) -> str:
    rows = t.rows
    if not t.header:
        rows = [[f"列{i + 1}" for i in range(len(rows[0]))], *rows]
    lines = ["| " + " | ".join(_cell(c) for c in rows[0]) + " |", "|" + " --- |" * len(rows[0])]
    lines += ["| " + " | ".join(_cell(c) for c in r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def _row_line(cells: list[str]) -> str:
    return "| " + " | ".join(_cell(c) for c in cells) + " |"


def iter_stream_table(t: StreamTable) -> Iterator[str]:
    """StreamTable を 1 行ずつ描画する（行の最後に改行は付けない）."""
    width = len(t.header)
    yield _row_line(t.header)
    yield "|" + " --- |" * width
    for r in t.rows:
        yield _row_line((list(r) + [""] * width)[:width])


def render_block(b: Block) -> str:
    """ListItem 以外の 1 ブロックを描画する."""
    if isinstance(b, Heading):
        return "#" * b.level + " " + b.text
    if isinstance(b, Paragraph):
        return "\\\n".join(b.text.split("\n"))  # 段落内の明示改行は GFM のハード改行
    if isinstance(b, Table):
        return render_table(b)
    if isinstance(b, Image):
        alt = b.alt.replace("[", "").replace("]", "")
        return f"![{alt}]({b.path})"
    if isinstance(b, Quote):
        lines = b.text.split("\n")
        head = f"**{b.label}**: " if b.label else ""
        return "\n".join(f"> {head if i == 0 else ''}{ln}" for i, ln in enumerate(lines))
    if isinstance(b, FootnoteDef):
        return f"[^{b.label}]: " + b.text.replace("\n", " ")
    if isinstance(b, Comment):
        return f"<!-- {b.text} -->"
    if isinstance(b, StreamTable):
        return "\n".join(iter_stream_table(b))
    raise TypeError(f"unknown block: {type(b).__name__}")


def iter_chunks(blocks: Iterable[Block]) -> Iterator[tuple[str, list[Block]]]:
    """(Markdown 断片, その断片を構成したブロック) を順に返す. 連続する ListItem は 1 断片にまとめる."""
    list_lines: list[str] = []
    list_blocks: list[Block] = []
    counters: dict[int, int] = {}
    for b in blocks:
        if isinstance(b, ListItem):
            counters = {lv: n for lv, n in counters.items() if lv <= b.level}
            counters[b.level] = counters.get(b.level, 0) + 1
            marker = f"{counters[b.level]}." if b.ordered else "-"
            list_lines.append("  " * b.level + f"{marker} {b.text}")
            list_blocks.append(b)
            continue
        if list_lines:
            yield "\n".join(list_lines), list(list_blocks)
            list_lines.clear()
            list_blocks.clear()
        counters = {}
        yield render_block(b), [b]
    if list_lines:
        yield "\n".join(list_lines), list(list_blocks)


def render(blocks: Iterable[Block]) -> str:
    chunks = [c for c, _ in iter_chunks(blocks)]
    return "\n\n".join(chunks).rstrip() + "\n"
