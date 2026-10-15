"""変換結果を AI に渡しやすい 1 つの索引・全文ファイルにまとめる（llms.txt 形式）.

- llms.txt      … 文書の一覧（タイトル・元ファイル・Markdown へのリンク）
- llms-full.txt … 全文書の Markdown を 1 ファイルに連結したもの（Front Matter は除く）

どちらも入力の並び（ファイル名順）だけで決まる。日時などは入れない。
"""

from __future__ import annotations

import re
from pathlib import Path

from .render import escape_comment, escape_inline, escape_markdown, link_destination

_FENCE = "---\n"


def _strip_front_matter(md: str) -> str:
    if md.startswith(_FENCE):
        end = md.find("\n---\n", len(_FENCE) - 1)
        if end >= 0:
            return md[end + len("\n---\n") :].lstrip("\n")
    return md


# エスケープされていない [ と ]（直前の \ が偶数個）
_UNESCAPED_BRACKET = re.compile(r"(?<!\\)((?:\\\\)*)([\[\]])")


def _link_text(markdown: str) -> str:
    """Markdown の文字列（見出しの行の中身）を、リンクの文字 [...] に入れられるようにする.

    見出しは描画で `\\`・`<`・`&` などをエスケープ済みだが、`[` `]` はそのまま（`# a]b(c`）.
    そのまま [..](..) に入れると `]` でリンクの文字が切れ、`](` でリンク先が変わるので、
    エスケープされていない `[` `]` の前に `\\` を付ける.
    """
    return _UNESCAPED_BRACKET.sub(r"\1\\\2", markdown)


def _plain(text: str) -> str:
    """素の文字列（ファイル名）を、Markdown の行の中に文字として書けるようにする."""
    return escape_inline(escape_markdown(text), markup=True)


def _title(md: str, fallback: str) -> str:
    """リンクの文字に使う、文書のタイトル（Markdown）. 見出しが無ければファイル名."""
    for line in _strip_front_matter(md).splitlines():
        if line.startswith("# "):
            return _link_text(line[2:].strip())
    return _plain(fallback)


def write(out_dir: Path, items: list[tuple[str, Path]], *, title: str = "変換済みドキュメント") -> list[Path]:
    """items: (元ファイル名, 出力 Markdown のパス) の並び. 書いたファイルを返す."""
    items = sorted(items, key=lambda x: x[1].name)
    index = [f"# {title}", "", f"> doc2md で Markdown に変換した {len(items)} 件の文書。", "", "## 文書", ""]
    full: list[str] = []
    for source, md_path in items:
        md = md_path.read_text(encoding="utf-8-sig")  # BOM 付きで保存し直されたファイルも読む
        name = _title(md, md_path.stem)
        rel = md_path.relative_to(out_dir).as_posix() if md_path.is_relative_to(out_dir) else md_path.as_posix()
        index.append(f"- [{name}]({link_destination(rel)}): 元ファイル {_plain(source)}")
        body = _strip_front_matter(md).rstrip("\n")
        full.append(f"<!-- source: {escape_comment(source)} -->\n\n{body}")
    out_dir.mkdir(parents=True, exist_ok=True)  # 全ファイルが失敗しても索引は作る
    index_path = out_dir / "llms.txt"
    full_path = out_dir / "llms-full.txt"
    index_path.write_text("\n".join(index) + "\n", encoding="utf-8", newline="\n")
    full_path.write_text("\n\n---\n\n".join(full) + "\n", encoding="utf-8", newline="\n")
    return [index_path, full_path]
