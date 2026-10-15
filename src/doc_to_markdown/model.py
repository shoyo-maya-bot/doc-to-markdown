"""形式に依存しない中間表現と変換オプション.

各アダプタは元文書を Block の列（リストまたはジェネレータ）に変換し、次の 2 種類の統計を添える。

- stats（SourceStats）: アダプタが「認識した」要素の数。正規化・描画で落ちたものの検知に使う
- raw_text_chars: アダプタの解釈とは **別の経路** で元ファイルから直接数えた本文の文字数
  （docx/pptx は XML の全テキスト、xlsx は全セル値、pdf は全文字）。アダプタ自体が見落とした
  テキスト（テキストボックス等）の検知に使う

大きなファイルでもメモリに載せきらないよう、blocks はジェネレータでもよい
（stats は blocks を最後まで読み終えた時点で確定する）。
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field


@dataclass
class Heading:
    level: int
    text: str
    synthetic: bool = False  # 変換側が付け足した見出し（ファイル名・区分名など）。ロス検知の対象外
    # text が変換側の書いた Markdown 記法（太字・リンク・脚注）を含む（Word 等）。
    # False（既定）は素の文字列として、描画で `\` `|` `<` 等をすべてエスケープする
    markup: bool = False


@dataclass
class Paragraph:
    text: str
    synthetic: bool = False
    markup: bool = False  # Heading.markup と同じ


@dataclass
class ListItem:
    text: str
    level: int = 0
    ordered: bool = False
    markup: bool = False  # Heading.markup と同じ
    # 番号付きのときの番号（元文書の番号付けで決まる続き番号. Word の List Number 等）.
    # None は描画側で連続する項目ごとに 1 から数える
    number: int | None = None


@dataclass
class Table:
    rows: list[list[str]]
    header: bool = True
    markup: bool = False  # セルが変換側の書いた Markdown 記法を含む（Heading.markup と同じ）


@dataclass
class StreamTable:
    """行をメモリに溜めずに書き出す表（大きな Excel 用）. 空列の削除などの正規化は行わない."""

    header: list[str]
    rows: Iterable[list[str]]


@dataclass
class Image:
    path: str  # Markdown からの相対パス
    alt: str = ""


@dataclass
class Quote:
    """本文の流れの外にあるテキスト（テキストボックス・図形内の文字など）. `> **label**: text` で出力."""

    text: str
    label: str = ""
    markup: bool = False  # Heading.markup と同じ


@dataclass
class FootnoteDef:
    """GFM の脚注定義 `[^label]: text`（脚注・文末脚注・コメント）."""

    label: str
    text: str


@dataclass
class Comment:
    """出典トレース用の HTML コメント（例: <!-- page 3 -->）。本文ではない。"""

    text: str


Block = Heading | Paragraph | ListItem | Table | StreamTable | Image | Quote | FootnoteDef | Comment


@dataclass
class SourceStats:
    headings: int = 0
    list_items: int = 0
    tables: int = 0
    table_rows: int = 0
    images: int = 0
    text_chars: int = 0  # アダプタが認識した本文の非空白文字数

    def add_text(self, text: str) -> None:
        self.text_chars += count_chars(text)


def count_chars(text: str) -> int:
    """空白以外の文字数（str.split は Unicode の空白＝全角スペース等も区切る）."""
    return len("".join(text.split()))


@dataclass
class Document:
    blocks: Iterable[Block] = field(default_factory=list)
    stats: SourceStats = field(default_factory=SourceStats)
    notes: list[str] = field(default_factory=list)  # 変換時の注意事項（レポートに info として出す）
    warnings: list[str] = field(default_factory=list)  # 利用者が確認すべき事項（レポートに warning として出す）
    # 変換で再現できなかった要素（rule_id, message, suggestion）. レポートに warning として出す
    # （例: loss.number_format）
    losses: list[tuple[str, str, str]] = field(default_factory=list)
    raw_text_chars: int | None = None  # 別経路で数えた元ファイルの本文文字数（None = 数えられない形式）


@dataclass
class Options:
    """変換オプション（CLI から渡す）. 既定値は「迷ったら情報を落とさない」側に寄せる."""

    via_pdf: bool = False
    # Word の変更履歴: accept = 承認後の姿（既定）/ reject = 変更前の姿 / show = 両方を記号付きで
    track_changes: str = "accept"
    include_comments: bool = True
    include_header_footer: bool = True
    # Excel
    header_rows: dict[str, int] = field(default_factory=dict)  # シート名 → 見出し行（1 始まり）。"*" は全シート
    skip_hidden: bool = False  # 非表示の行・列・シートを出力しない
    recalc: bool = False  # 数式の計算結果が無いブックを LibreOffice で再計算してから読む
    # これを超えるセル数の xlsx は省メモリ読み込み（セル結合・複数表・非表示の判定なし）。
    # 通常読み込みはセル 100 万個で約 800MB 使うため、既定は 20 万セル（約 150MB）
    xlsx_full_load_max_cells: int = 200_000
    # Office 形式（zip）の展開の上限（zip 爆弾対策）. 超えたら変換せずに止める（終了コード 2）
    max_unzipped_mb: int = 1024  # 展開後の合計サイズ（MB）
    max_zip_ratio: int = 500  # 1 MB 以上のファイルの圧縮率（展開後 / 圧縮後）
    # PDF
    ocr: str = "auto"  # auto = 文字の無いページだけ / always / never
    ocr_lang: str = "jpn+eng"
    ocr_dpi: int = 300
    remove_running_headers: bool = True  # 全ページに繰り返し出る行（ヘッダー・フッター・ページ番号）を除く
    # ロス検知（loss.text / loss.raw_text）のしきい値: 出力の本文文字数が元の何 % 未満なら警告するか.
    # 文字数だけを見るので、短い語の欠落・語の入れ替わり・誤った文字への置き換えは検知できない
    loss_threshold: float = 95.0
    # 出力 Markdown の先頭に YAML の Front Matter（title / id / source / source_sha256 / generator）を付ける
    front_matter: bool = False
