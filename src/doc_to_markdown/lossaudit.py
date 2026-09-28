"""変換ロス検知: 元文書側で数えた統計と、出力 Markdown から数えた統計を突き合わせる.

- 見出し数・箇条書き数・表の数・表の行数・画像数: 一致しなければ warning
- 本文文字数（非空白文字）: 出力側が元の 95% 未満なら warning（表記の正規化による差は許容）
"""

from __future__ import annotations

import re

from devtools_common.report import Finding, Severity

from .model import SourceStats

_HEADING = re.compile(r"^#{1,6} ")
_LIST = re.compile(r"^\s*(?:-|\d+\.) ")
_TABLE_SEP = re.compile(r"^\|(?: --- \|)+$")
_IMAGE = re.compile(r"!\[[^\]]*\]\([^)]+\)")
_COMMENT = re.compile(r"<!--.*?-->")
_LINK = re.compile(r"\[([^\]]*)\]\([^)]+\)")
_MARKUP = re.compile(r"(\*\*\*|\*\*|\*|\\\||\\\\|<br>|^#{1,6} |^\s*(?:-|\d+\.) |\|)", re.M)

TEXT_RATIO_MIN = 0.95


def markdown_stats(md: str) -> SourceStats:
    st = SourceStats()
    lines = md.splitlines()
    in_table = False
    for line in lines:
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
    body = _COMMENT.sub("", md)
    body = _IMAGE.sub("", body)
    body = _LINK.sub(r"\1", body)
    body = _MARKUP.sub("", body)
    body = body.replace("\\\n", "\n")
    st.add_text(body)
    return st


def audit(source: SourceStats, md: str, *, location: str) -> list[Finding]:
    """md は変換側が付け足した要素（synthetic）を除いて描画した Markdown を渡す."""
    out = markdown_stats(md)
    findings: list[Finding] = []
    for name, label in [
        ("headings", "見出し数"),
        ("list_items", "箇条書き数"),
        ("tables", "表の数"),
        ("table_rows", "表の行数"),
        ("images", "画像数"),
    ]:
        src, dst = getattr(source, name), getattr(out, name)
        if src != dst:
            findings.append(
                Finding(
                    rule_id=f"loss.{name}",
                    severity=Severity.WARNING,
                    location=location,
                    field=name,
                    message=f"{label}が元文書と一致しません",
                    expected=str(src),
                    actual=str(dst),
                    suggestion="--via-pdf での変換結果と比較してください" if dst < src else "",
                )
            )
    if source.text_chars:
        ratio = out.text_chars / source.text_chars
        if ratio < TEXT_RATIO_MIN:
            findings.append(
                Finding(
                    rule_id="loss.text",
                    severity=Severity.WARNING,
                    location=location,
                    field="text_chars",
                    message=f"本文の文字数が元文書の {ratio:.0%} です（欠落の可能性）",
                    expected=str(source.text_chars),
                    actual=str(out.text_chars),
                    suggestion="欠落箇所を目視確認し、必要なら --via-pdf を試してください",
                )
            )
    return findings
