"""中間表現 → Markdown（GFM）. 同じ入力からは常に同じ文字列を返す."""

from __future__ import annotations

from .model import Block, Comment, Heading, Image, ListItem, Paragraph, Table


def _cell(text: str) -> str:
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _table(t: Table) -> str:
    rows = t.rows
    if not t.header:
        rows = [[f"列{i + 1}" for i in range(len(rows[0]))], *rows]
    lines = ["| " + " | ".join(_cell(c) for c in rows[0]) + " |", "|" + " --- |" * len(rows[0])]
    lines += ["| " + " | ".join(_cell(c) for c in r) + " |" for r in rows[1:]]
    return "\n".join(lines)


def render(blocks: list[Block]) -> str:
    chunks: list[str] = []
    list_lines: list[str] = []
    counters: dict[int, int] = {}

    def flush_list() -> None:
        nonlocal counters
        if list_lines:
            chunks.append("\n".join(list_lines))
            list_lines.clear()
        counters = {}

    for b in blocks:
        if isinstance(b, ListItem):
            counters = {lv: n for lv, n in counters.items() if lv <= b.level}
            counters[b.level] = counters.get(b.level, 0) + 1
            marker = f"{counters[b.level]}." if b.ordered else "-"
            list_lines.append("  " * b.level + f"{marker} {b.text}")
            continue
        flush_list()
        if isinstance(b, Heading):
            chunks.append("#" * b.level + " " + b.text)
        elif isinstance(b, Paragraph):
            chunks.append("\\\n".join(b.text.split("\n")))  # 段落内の明示改行は GFM のハード改行
        elif isinstance(b, Table):
            chunks.append(_table(b))
        elif isinstance(b, Image):
            alt = b.alt.replace("[", "").replace("]", "")
            chunks.append(f"![{alt}]({b.path})")
        elif isinstance(b, Comment):
            chunks.append(f"<!-- {b.text} -->")
    flush_list()
    return "\n\n".join(chunks).rstrip() + "\n"
