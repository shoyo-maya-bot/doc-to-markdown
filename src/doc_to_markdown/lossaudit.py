"""変換ロス検知.

2 段構えで検知する。

1. 構造の一致: アダプタが認識した要素数（SourceStats）と、出力 Markdown から数えた要素数を比べる。
   → 正規化・描画の段階で落ちたものを検知する（見出し・箇条書き・表・行・画像）
2. 本文量の一致: 元ファイルから **アダプタとは別の経路** で数えた本文文字数（raw_text_chars）と、
   出力の本文文字数を比べる。→ アダプタ自体が見落としたテキスト（未対応の要素）を検知する

どちらも「出力が少ない」方向だけを問題にする（表の縦結合の値を繰り返す等で出力が増えるのは正常）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from devtools_common.report import Finding, Severity

from .model import SourceStats, count_chars

_HEADING = re.compile(r"^#{1,6} ")
_LIST = re.compile(r"^\s*(?:-|\d+\.) ")
_TABLE_SEP = re.compile(r"^\|(?: --- \|)+$")
_IMAGE = re.compile(r"!\[[^\]]*\]\([^)]+\)")
_COMMENT = re.compile(r"(?<!\\)<!--.*?-->", re.S)  # `\<!--` はエスケープされた本文
_LINK = re.compile(r"\[([^\]]*)\]\([^)]+\)")
_FOOTNOTE = re.compile(r"(?<!\\)\[\^[^\]]+\]:?")
_QUOTE_LABEL = re.compile(r"^> (\*\*[^*]+\*\*: )?", re.M)
_MARKUP = re.compile(r"(\*\*\*|\*\*|\*|~~|</?ins>|\\\||<br>|^#{1,6} |^\s*(?:-|\d+\.) |\|)", re.M)
# バックスラッシュによるエスケープ（`\#` `\<` `\\` 等）は元の 1 文字として数える
_ESCAPED = re.compile(r"\\([!-/:-@\[-`{-~])")

TEXT_RATIO_MIN = 0.95  # 既定のしきい値（--loss-threshold 95）
STRUCTURE_FIELDS = [
    ("headings", "見出し数"),
    ("list_items", "箇条書き数"),
    ("tables", "表の数"),
    ("table_rows", "表の行数"),
    ("images", "画像数"),
]


def markdown_stats(md: str, st: SourceStats | None = None) -> SourceStats:
    st = st or SourceStats()
    in_table = False
    for line in md.splitlines():
        if _HEADING.match(line):
            st.headings += 1
        elif _LIST.match(line):
            st.list_items += 1
        if _TABLE_SEP.match(line):
            st.tables += 1
            in_table = True
            st.table_rows += 1  # ヘッダ行（直前の行）
            continue
        if in_table:
            if line.startswith("|"):
                st.table_rows += 1
            else:
                in_table = False
        st.images += len(_IMAGE.findall(line))
    st.text_chars += count_chars(plain_text(md))
    return st


_PUA = 0xE000  # エスケープされた記号を、記法の除去が終わるまで私用領域の文字に退避する


def _keep_escaped(m: re.Match[str]) -> str:
    # `\|` は表の区切りと同じく数えない（従来どおり）。それ以外は記号 1 文字として残す
    return "\\|" if m.group(1) == "|" else chr(_PUA + ord(m.group(1)))


def _restore_escaped(text: str) -> str:
    return re.sub("[\ue021-\ue07e]", lambda m: chr(ord(m.group(0)) - _PUA), text)


def plain_text(md: str) -> str:
    # エスケープした記号（`\[` `\*` 等）を先に退避してから記法を外す.
    # 後にすると、本文の `\[x\](…)` をリンクとみなして文字を落とし、ロスと誤検知する
    body = _ESCAPED.sub(_keep_escaped, md)
    body = _COMMENT.sub("", body)
    body = _IMAGE.sub("", body)
    body = _LINK.sub(r"\1", body)
    body = _FOOTNOTE.sub("", body)
    body = _QUOTE_LABEL.sub("", body)
    body = _restore_escaped(_MARKUP.sub("", body))
    return body.replace("\\\n", "\n")


@dataclass
class OutputStats:
    """出力 Markdown の統計を断片ごとに積み上げる（synthetic なブロックは除く）."""

    structure: SourceStats = field(default_factory=SourceStats)

    def add(self, markdown: str, *, synthetic: bool) -> None:
        if not synthetic:
            markdown_stats(markdown, self.structure)


def audit(
    source: SourceStats,
    out: OutputStats,
    *,
    location: str,
    raw_text_chars: int | None = None,
    text_ratio_min: float = TEXT_RATIO_MIN,
) -> list[Finding]:
    """ロスの疑いを返す. text_ratio_min は本文文字数の比のしきい値（0〜1. 既定 0.95）.

    文字数の比だけを見るので、しきい値より小さな欠落（長い文書の中の 1 文・1 語）、
    語の入れ替わり、別の文字への置き換え（文字化け）は検知できない。
    """
    findings: list[Finding] = []
    for name, label in STRUCTURE_FIELDS:
        src, dst = getattr(source, name), getattr(out.structure, name)
        if dst < src:
            findings.append(
                Finding(
                    rule_id=f"loss.{name}",
                    severity=Severity.WARNING,
                    location=location,
                    field=name,
                    message=f"{label}が元文書より少なくなっています",
                    expected=str(src),
                    actual=str(dst),
                    suggestion="--via-pdf での変換結果と比較してください",
                )
            )
    if source.text_chars and out.structure.text_chars < source.text_chars * text_ratio_min:
        ratio = out.structure.text_chars / source.text_chars
        findings.append(
            Finding(
                rule_id="loss.text",
                severity=Severity.WARNING,
                location=location,
                field="text_chars",
                message=f"本文の文字数が、変換時に認識した量の {ratio:.0%} です（整形で欠落した可能性）",
                expected=str(source.text_chars),
                actual=str(out.structure.text_chars),
                suggestion="欠落箇所を目視確認してください",
            )
        )
    if raw_text_chars:
        produced = out.structure.text_chars
        if produced < raw_text_chars * text_ratio_min:
            ratio = produced / raw_text_chars
            findings.append(
                Finding(
                    rule_id="loss.raw_text",
                    severity=Severity.WARNING,
                    location=location,
                    field="raw_text_chars",
                    message=f"元ファイルの文字の {ratio:.0%} しか出力されていません（未対応の要素がある可能性）",
                    expected=str(raw_text_chars),
                    actual=str(produced),
                    suggestion="図形・グラフ・埋め込みオブジェクト等の未対応要素を確認し、必要なら --via-pdf を試す",
                )
            )
    return findings
