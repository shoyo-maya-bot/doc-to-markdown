"""draw.io (.drawio) → 中間表現.

- 図は svg に書き出して Markdown から参照（draw.io デスクトップの CLI がある場合）
- 編集可能ソース（.drawio）も同じ assets/ にコピーし、書き出しとペアで保管する
- 図の中身（図形ラベルと接続）をテキストでも出力し、検索・AI 読解・差分レビューに使えるようにする
- 圧縮形式（deflate + base64）の diagram にも対応（展開後のサイズは全ページの合計で --max-unzipped-mb まで）
- XML は defusedxml で解析する（信頼できない入力を想定）
"""

from __future__ import annotations

import base64
import html
import re
import subprocess
import urllib.parse
import xml.etree.ElementTree as ET  # noqa: S405 — 型と空要素の生成のみ。解析は defusedxml
import zlib
from pathlib import Path

import defusedxml.ElementTree as SafeET
from devtools_common.cli import UsageError
from devtools_common.executables import DRAWIO, find_executable

from ..assets import AssetWriter
from ..model import Document, Heading, Image, ListItem, Options, Paragraph
from ..render import escape_markdown, link_destination

_TAG = re.compile(r"<[^>]+>")


def _label(value: str | None) -> str:
    if not value:
        return ""
    text = _TAG.sub(" ", value.replace("<br>", "\n").replace("<br/>", "\n"))
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def _inflate(data: bytes, limit: int, name: str) -> bytes:
    """deflate（raw）を展開する. 展開後が limit バイトを超えたら途中で止める（圧縮爆弾対策）."""
    d = zlib.decompressobj(-15)
    out = d.decompress(data, limit + 1)
    if len(out) > limit or d.unconsumed_tail:
        raise UsageError(
            f"{name} の圧縮された図の展開後のサイズが上限 {limit // (1024 * 1024):,} MB を超えるため変換しません"
            "（壊れたファイルか圧縮爆弾の疑い。正しいファイルなら --max-unzipped-mb で上限を上げてください）"
        )
    return out + d.flush()


class InflateBudget:
    """1 ファイルの中の圧縮された図の、展開後の合計サイズの残り（ページごとの上限だけでは、
    上限ぎりぎりのページを多数並べた図で合計が際限なく増える）."""

    def __init__(self, limit: int):
        self.limit = limit
        self.remaining = limit

    def take(self, data: bytes, name: str) -> bytes:
        try:
            out = _inflate(data, self.remaining, name)
        except UsageError:
            raise UsageError(
                f"{name} の圧縮された図の展開後のサイズ（全ページの合計）が上限"
                f" {self.limit // (1024 * 1024):,} MB を超えるため変換しません"
                "（壊れたファイルか圧縮爆弾の疑い。正しいファイルなら --max-unzipped-mb で上限を上げてください）"
            ) from None
        self.remaining -= len(out)
        return out


def _diagram_root(
    diagram: ET.Element, *, limit: int = 1024 << 20, name: str = "", budget: InflateBudget | None = None
) -> ET.Element | None:
    model = diagram.find("mxGraphModel")
    if model is not None:
        return model.find("root")
    raw = (diagram.text or "").strip()
    if not raw:
        return None
    budget = budget or InflateBudget(limit)
    xml = urllib.parse.unquote(budget.take(base64.b64decode(raw), name).decode("utf-8"))
    return SafeET.fromstring(xml).find("root")


def find_drawio_cli() -> str | None:
    return find_executable(DRAWIO)


def export_svg(src: Path, dest: Path, cli: str) -> bool:
    try:
        subprocess.run(
            [cli, "--export", "--format", "svg", "--output", str(dest), str(src)],
            check=True,
            capture_output=True,
            timeout=120,
        )
    except (subprocess.SubprocessError, OSError):
        return False
    return dest.is_file()


def convert(path: Path, assets: AssetWriter, opts: Options | None = None) -> Document:
    opts = opts or Options()
    limit = opts.max_unzipped_mb * 1024 * 1024  # 圧縮された図の展開後の上限（Office の zip と同じ値を使う）
    doc = Document()
    st = doc.stats
    doc.blocks.append(Heading(1, path.stem, synthetic=True))

    source_rel = assets.copy_file(path, path.name)
    source_link = f"[{escape_markdown(path.name)}]({link_destination(source_rel)})"  # 空白・括弧を含む名前でも切れない
    cli = find_drawio_cli()
    svg_name = f"{path.stem}.svg"
    if cli and export_svg(path, assets.dir / svg_name, cli):
        st.images += 1
        doc.blocks.append(Image(f"{assets.dir_name}/{svg_name}", alt=path.stem))
        doc.blocks.append(Paragraph(f"編集用ソース: {source_link}", markup=True))
    else:
        doc.notes.append(
            "draw.io Desktop が見つからないため svg 書き出しを省略（ソースのみ保存。場所は DEVTOOLS_DRAWIO で指定可）"
        )
        doc.blocks.append(Paragraph(f"図のソース: {source_link}（svg は draw.io で書き出してください）", markup=True))

    tree = SafeET.parse(path)  # 信頼できない XML（外部実体・展開攻撃）対策
    budget = InflateBudget(limit)  # 上限はページごとではなくファイル全体の合計で見る
    for page_no, diagram in enumerate(tree.getroot().iter("diagram"), 1):
        root = _diagram_root(diagram, name=path.name, budget=budget)
        name = diagram.get("name") or f"Page-{page_no}"
        st.headings += 1
        doc.blocks.append(Heading(2, name))
        if root is None:
            continue
        cells = {c.get("id"): c for c in root.iter("mxCell")}
        vertices = [c for c in cells.values() if c.get("vertex") == "1" and _label(c.get("value"))]
        edges = [c for c in cells.values() if c.get("edge") == "1"]
        if vertices:
            doc.blocks.append(Heading(3, "要素", synthetic=True))
            for v in vertices:
                text = _label(v.get("value"))
                st.add_text(text)
                st.list_items += 1
                doc.blocks.append(ListItem(text))
        if edges:
            doc.blocks.append(Heading(3, "接続", synthetic=True))
            for e in edges:
                src = _label(cells.get(e.get("source"), ET.Element("x")).get("value")) or "?"
                dst = _label(cells.get(e.get("target"), ET.Element("x")).get("value")) or "?"
                label = _label(e.get("value"))
                if label:
                    st.add_text(label)
                st.list_items += 1
                doc.blocks.append(ListItem(f"{src} → {dst}" + (f"（{label}）" if label else "")))
    return doc
