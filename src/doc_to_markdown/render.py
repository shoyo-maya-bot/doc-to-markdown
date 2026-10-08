"""中間表現 → Markdown（GFM）. 同じ入力からは常に同じ文字列を返す.

`iter_chunks` はブロックを順に Markdown の断片（段落単位）にして返すので、大きな文書も
ファイルへ逐次書き出せる。断片どうしは空行 1 つで区切る。
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Iterator

from .model import Block, Comment, FootnoteDef, Heading, Image, ListItem, Paragraph, Quote, StreamTable, Table

# 文字列の中の HTML（タグ・コメント）。描画で消えたり解釈されたりしないよう `\<` にする.
# Word の変更履歴（<ins>）は変換側が付ける記法なので残す
_HTML = re.compile(r"<(?!/?ins>)(?=[A-Za-z/!?])")
# `&lt;` などの文字参照（そのまま書くと `<` に化ける）
_ENTITY = re.compile(r"&(?=#?[0-9A-Za-z]+;)")
# 行頭にあると Markdown のブロック記法として解釈されるもの
_LINE_START = re.compile(
    r"""^(?:
        \#{1,6}(?:[ \t]|$)          # 見出し
      | >                           # 引用
      | [-+*](?:[ \t]|$)            # 箇条書き
      | [-*_=][-*_= \t]*$           # 区切り線・Setext 見出しの下線
      | `{3,} | ~{3,}               # コードブロック
      | \[[^\]]*\]:                 # リンク参照の定義・脚注の定義（行ごと消える）
    )""",
    re.X,
)
_ORDERED = re.compile(r"^(\d{1,9})([.)])(?=[ \t]|$)")
# 見出しの末尾の # は閉じ記号として消える
_CLOSING_HASHES = re.compile(r"([ \t])(#+)[ \t]*$")


# 素の文字列のうち、行の途中でも Markdown のインライン記法（強調・コード・リンク・打ち消し）になる記号
_INLINE_MARKUP = re.compile(r"([\\`*_\[\]~])")
# リンク先に入れるとリンクが途中で切れる・解釈が変わる文字（空白・括弧・山括弧・制御文字・`\\`）.
# URL の `%` は既にエスケープ（%20 等）でなければ %25 にする. ファイルのパスの `%` は常に %25 にする
_DEST_UNSAFE_URL = re.compile(r"[\x00-\x20\x7f()<>\\]|%(?![0-9A-Fa-f]{2})")
_DEST_UNSAFE_PATH = re.compile(r"[\x00-\x20\x7f()<>\\%]")


def escape_markdown(text: str) -> str:
    """素の文字列を、Markdown 記法を含む文字列（markup=True のブロック）に埋め込むためにエスケープする.

    `\\` と、強調・コード・リンク・打ち消しになる記号（`` ` `` `*` `_` `[` `]` `~`）の前に `\\` を付ける。
    行頭のブロック記法・HTML・文字参照は描画（escape_text 等）が別にエスケープする。
    """
    return _INLINE_MARKUP.sub(r"\\\1", text)


def _percent(m: re.Match[str]) -> str:
    return "".join(f"%{b:02X}" for b in m.group(0).encode("utf-8"))


def link_destination(dest: str, *, path: bool = True) -> str:
    """リンク・画像のリンク先を、空白や括弧を含んでも壊れない形にする（パーセントエンコード）.

    `[文字](a (1).png)` は `)` でリンクが切れ、空白を含むとリンクにならない。該当する文字だけを
    `%20` `%28` `%29` 等にする（日本語などはそのまま残す）。
    path=True（既定）は出力フォルダからの相対パス（`%` も文字としてエンコードする）。
    path=False は文書に書かれた URL（`%20` のように既にエンコード済みの部分はそのまま残す）。
    """
    return (_DEST_UNSAFE_PATH if path else _DEST_UNSAFE_URL).sub(_percent, dest)


def escape_inline(text: str, *, markup: bool = False) -> str:
    """文字列の中の Markdown・HTML として解釈される文字をエスケープする.

    markup=True は変換側が意図して書いた記法（太字・リンク・脚注）を含む文字列: `\\` と `|` は触らない。
    その中の素の文字列は、作る側（アダプタ）が escape_markdown でエスケープしておく。
    """
    if not markup:
        text = text.replace("\\", "\\\\").replace("|", "\\|")
    return _ENTITY.sub("\\\\&", _HTML.sub("\\\\<", text))


def escape_line_start(line: str) -> str:
    """行頭のブロック記法（見出し・箇条書き・番号付き・引用・区切り線 等）をエスケープする.

    行頭の空白は Markdown では意味を持たない（4 つ以上はコードブロックになる）ので取り除く。
    """
    line = line.lstrip(" \t")
    if _LINE_START.match(line):
        return "\\" + line
    return _ORDERED.sub(r"\1\\\2", line, count=1)


def escape_text(text: str, *, markup: bool = False) -> list[str]:
    """段落の本文を行ごとにエスケープする（段落の途中の行もブロック記法になり得る）."""
    return [escape_line_start(escape_inline(ln, markup=markup)) for ln in text.split("\n")]


def escape_comment(text: str) -> str:
    """HTML コメントの中身: `--`（`-->` でコメントが途中で閉じる）を崩し、1 行にする."""
    text = text.replace("\n", " ")
    while "--" in text:
        text = text.replace("--", "- -")
    return text


def _list_line(b: ListItem, marker: str) -> str:
    """箇条書き 1 行. 本文の行頭が見出し・引用などに化けないよう、段落と同じくエスケープする."""
    text = escape_line_start(escape_inline(b.text, markup=b.markup).replace("\n", " "))
    return "  " * b.level + f"{marker} {text}"


def _cell(text: str, markup: bool = False) -> str:
    if markup:  # 記法を含むセル（Word）: `\\` はアダプタがエスケープ済み. 列の区切りになる `|` だけを足す
        return escape_inline(text, markup=True).replace("|", "\\|").replace("\n", "<br>")
    return escape_inline(text).replace("\n", "<br>")


def render_table(t: Table) -> str:
    rows = t.rows
    if not t.header:
        rows = [[f"列{i + 1}" for i in range(len(rows[0]))], *rows]
    m = t.markup
    lines = ["| " + " | ".join(_cell(c, m) for c in rows[0]) + " |", "|" + " --- |" * len(rows[0])]
    lines += ["| " + " | ".join(_cell(c, m) for c in r) + " |" for r in rows[1:]]
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
        text = escape_inline(b.text, markup=b.markup).replace("\n", " ").strip()
        return "#" * b.level + " " + _CLOSING_HASHES.sub(r"\1\\\2", text)
    if isinstance(b, Paragraph):
        return "\\\n".join(escape_text(b.text, markup=b.markup))  # 段落内の明示改行は GFM のハード改行
    if isinstance(b, Table):
        return render_table(b)
    if isinstance(b, Image):
        alt = b.alt.replace("[", "").replace("]", "")
        return f"![{alt}]({link_destination(b.path)})"
    if isinstance(b, Quote):
        lines = escape_text(b.text, markup=b.markup)
        head = f"**{escape_inline(b.label)}**: " if b.label else ""
        return "\n".join(f"> {head if i == 0 else ''}{ln}" for i, ln in enumerate(lines))
    if isinstance(b, FootnoteDef):
        return f"[^{b.label}]: " + b.text.replace("\n", " ")
    if isinstance(b, Comment):
        return f"<!-- {escape_comment(b.text)} -->"
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
            list_lines.append(_list_line(b, marker))
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
