"""変換の流れ: (前処理) → アダプタ → 正規化 → 描画（逐次書き出し）→ ロス検知.

大きな文書でもメモリに載せきらないよう、アダプタが返すブロックを 1 つずつ正規化・描画し、
Markdown ファイルへ順に書き出す。ロス検知用の統計も書き出しながら積み上げる。
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path

from devtools_common.cli import UsageError
from devtools_common.log import get_logger
from devtools_common.report import Finding, Severity
from devtools_common.tempfiles import secure_tempdir

from . import lossaudit
from .adapters import ADAPTERS, OFFICE, SUPPORTED, VIA_PDF_ONLY, pdf_adapter
from .assets import AssetWriter
from .model import Document, Heading, ListItem, Options, SourceStats, StreamTable, count_chars
from .normalize import normalize
from .office import find_soffice, libreoffice_convert, office_to_pdf
from .render import iter_chunks, iter_stream_table

log = get_logger(__name__)


@dataclass
class Result:
    source: Path
    output: Path
    stats: SourceStats
    assets: int
    via_pdf: bool
    findings: list[Finding] = field(default_factory=list)


def _write(doc: Document, output: Path) -> lossaudit.OutputStats:
    stats = lossaudit.OutputStats()
    first = True
    with output.open("w", encoding="utf-8", newline="\n") as f:
        for b in _stream_aware(normalize(doc.blocks)):
            if isinstance(b, StreamTable):  # 大きな表は 1 行ずつ書く
                if not first:
                    f.write("\n\n")
                for i, line in enumerate(iter_stream_table(b)):
                    f.write(("\n" if i else "") + line)
                    if i == 1:  # 区切り行: 表 1 つ + 見出し行 1 行
                        stats.structure.tables += 1
                        stats.structure.table_rows += 1
                        continue
                    if i > 1:
                        stats.structure.table_rows += 1
                    stats.structure.text_chars += count_chars(lossaudit.plain_text(line))
                first = False
                continue
            chunk, blocks = b
            if not first:
                f.write("\n\n")
            f.write(chunk)
            first = False
            stats.add(chunk, synthetic=all(getattr(x, "synthetic", False) for x in blocks))
        f.write("\n")
    return stats


def _stream_aware(blocks):
    """StreamTable はそのまま、それ以外は (Markdown 断片, ブロック列) にして返す."""
    pending = []

    def flush():
        if pending:
            yield from iter_chunks(list(pending))
            pending.clear()

    for b in blocks:
        if isinstance(b, StreamTable):
            yield from flush()
            yield b
        elif isinstance(b, ListItem):
            pending.append(b)  # 連続する箇条書きは 1 断片にまとめる
        else:
            yield from flush()
            yield from iter_chunks([b])
    yield from flush()


def convert_file(src: Path, out_dir: Path, *, via_pdf: bool = False, options: Options | None = None) -> Result:
    opts = options or Options(via_pdf=via_pdf)
    if via_pdf:
        opts.via_pdf = True
    suffix = src.suffix.lower()
    if suffix not in SUPPORTED:
        raise UsageError(f"未対応の形式です: {src.name}（対応: {', '.join(sorted(SUPPORTED))}）")
    use_pdf = suffix in VIA_PDF_ONLY or (opts.via_pdf and suffix in OFFICE)
    stem = src.stem
    out_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = out_dir / f"{stem}.assets"
    if assets_dir.exists():  # 再実行時に古い抽出図が残らないようにする（決定的な出力）
        shutil.rmtree(assets_dir)
    assets = AssetWriter(out_dir, stem)
    output = out_dir / f"{stem}.md"

    if use_pdf:
        with secure_tempdir("doc2md-") as tmp:
            pdf = office_to_pdf(src, tmp)
            doc = pdf_adapter.convert(pdf, assets, opts)
            # PDF のファイル名由来の H1 を元ファイル名に戻す（ブロックは逐次生成なので先頭だけ差し替える）
            doc.blocks = _retitle(doc.blocks, stem)
            out = _write(doc, output)  # 一時 PDF を読み終えてから削除する
    else:
        doc = ADAPTERS[suffix](src, assets, opts)
        out = _write(doc, output)

    findings = lossaudit.audit(doc.stats, out, location=src.name, raw_text_chars=doc.raw_text_chars)
    findings += [Finding("convert.warning", Severity.WARNING, w, location=src.name) for w in doc.warnings]
    findings += [Finding("convert.note", Severity.INFO, n, location=src.name) for n in doc.notes]
    log.info("converted %s -> %s (%d chars, %d assets)", src.name, output.name, doc.stats.text_chars, assets.count)
    return Result(src, output, doc.stats, assets.count, use_pdf, findings)


def _retitle(blocks, stem: str):
    first = True
    for b in blocks:
        if first and isinstance(b, Heading) and b.synthetic:
            b = Heading(1, stem, synthetic=True)
        first = False
        yield b


def collect_inputs(paths: list[Path], *, recursive: bool) -> list[Path]:
    files: list[Path] = []
    for p in paths:
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            files += [f for f in it if f.is_file() and f.suffix.lower() in SUPPORTED and not f.name.startswith("~$")]
        elif p.is_file():
            files.append(p)
        else:
            raise UsageError(f"入力が見つかりません: {p}")
    return sorted(set(files))


__all__ = [
    "Result",
    "convert_file",
    "collect_inputs",
    "find_soffice",
    "office_to_pdf",
    "libreoffice_convert",
    "Options",
]
