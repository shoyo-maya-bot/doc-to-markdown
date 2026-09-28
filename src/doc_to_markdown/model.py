"""形式に依存しない中間表現.

各アダプタは元文書を Block 列に変換し、元文書側で数えた統計（SourceStats）を添える。
統計は変換ロス検知（lossaudit）で、出力 Markdown から数えた値と突き合わせる。
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Heading:
    level: int
    text: str
    synthetic: bool = False  # 変換側が付け足した見出し（ファイル名・区分名など）。ロス検知の対象外


@dataclass
class Paragraph:
    text: str
    synthetic: bool = False


@dataclass
class ListItem:
    text: str
    level: int = 0
    ordered: bool = False


@dataclass
class Table:
    rows: list[list[str]]
    header: bool = True


@dataclass
class Image:
    path: str  # Markdown からの相対パス
    alt: str = ""


@dataclass
class Comment:
    """出典トレース用の HTML コメント（例: <!-- page 3 -->）。本文ではない。"""

    text: str


Block = Heading | Paragraph | ListItem | Table | Image | Comment


@dataclass
class SourceStats:
    headings: int = 0
    list_items: int = 0
    tables: int = 0
    table_rows: int = 0
    images: int = 0
    text_chars: int = 0  # 元文書の本文の非空白文字数

    def add_text(self, text: str) -> None:
        self.text_chars += sum(1 for ch in text if not ch.isspace())


@dataclass
class Document:
    blocks: list[Block] = field(default_factory=list)
    stats: SourceStats = field(default_factory=SourceStats)
    notes: list[str] = field(default_factory=list)  # 変換時の注意事項（レポートに info として出す）
