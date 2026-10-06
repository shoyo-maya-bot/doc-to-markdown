"""変換の流れ: (前処理) → アダプタ → 正規化 → 描画（逐次書き出し）→ ロス検知.

大きな文書でもメモリに載せきらないよう、アダプタが返すブロックを 1 つずつ正規化・描画し、
Markdown ファイルへ順に書き出す。ロス検知用の統計も書き出しながら積み上げる。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
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


def _write(doc: Document, output: Path, header: str = "") -> lossaudit.OutputStats:
    stats = lossaudit.OutputStats()
    first = True
    with output.open("w", encoding="utf-8", newline="\n") as f:
        f.write(header)  # Front Matter（ロス検知の統計には含めない）
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


def convert_file(
    src: Path,
    out_dir: Path,
    *,
    via_pdf: bool = False,
    options: Options | None = None,
    name: str | None = None,
) -> Result:
    """src を変換して `<out_dir>/<name>.md` と `<out_dir>/<name>.assets/` に書き出す.

    name を省略すると元ファイル名の拡張子を除いたもの。複数ファイルを 1 つの出力先に書くときは
    `output_names` で重ならない名前を決めてから渡す（見出しや Front Matter の title は常に元ファイル名）。
    """
    opts = options or Options(via_pdf=via_pdf)
    if via_pdf:
        opts.via_pdf = True
    suffix = src.suffix.lower()
    if suffix not in SUPPORTED:
        raise UsageError(f"未対応の形式です: {src.name}（対応: {', '.join(sorted(SUPPORTED))}）")
    use_pdf = suffix in VIA_PDF_ONLY or (opts.via_pdf and suffix in OFFICE)
    stem = src.stem
    name = name or stem
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / f"{name}.md"
    findings: list[Finding] = []
    prev_source = existing_source(output)
    other_source = prev_source not in ("", src.name)
    if other_source:  # 前回の実行で別の元ファイルから作った出力（Front Matter の source で判定）
        log.warning("overwriting %s (previously converted from %s) with %s", output.name, prev_source, src.name)
        findings.append(
            Finding(
                "convert.overwrite_other_source",
                Severity.WARNING,
                f"既存の {output.name} は別の元ファイル {prev_source} から作られたものですが、上書きしました",
                location=src.name,
                suggestion="出力先を分けるか、不要になった古い出力を削除してください",
            )
        )
    assets_dir = out_dir / f"{name}.assets"
    if assets_dir.exists():  # 再実行時に古い抽出図が残らないようにする（決定的な出力）
        shutil.rmtree(assets_dir)
    assets = AssetWriter(out_dir, name)
    doc_id = "" if other_source else existing_id(output)  # 別文書の id は引き継がない
    header = front_matter(src, stem, doc_id) if opts.front_matter else ""

    if use_pdf:
        with secure_tempdir("doc2md-") as tmp:
            pdf = office_to_pdf(src, tmp)
            doc = pdf_adapter.convert(pdf, assets, opts)
            # PDF のファイル名由来の H1 を元ファイル名に戻す（ブロックは逐次生成なので先頭だけ差し替える）
            doc.blocks = _retitle(doc.blocks, stem)
            out = _write(doc, output, header)  # 一時 PDF を読み終えてから削除する
    else:
        doc = ADAPTERS[suffix](src, assets, opts)
        out = _write(doc, output, header)

    findings += lossaudit.audit(doc.stats, out, location=src.name, raw_text_chars=doc.raw_text_chars)
    findings += [Finding("convert.warning", Severity.WARNING, w, location=src.name) for w in doc.warnings]
    findings += [
        Finding(rule, Severity.WARNING, msg, location=src.name, suggestion=hint) for rule, msg, hint in doc.losses
    ]
    findings += [Finding("convert.note", Severity.INFO, n, location=src.name) for n in doc.notes]
    log.info("converted %s -> %s (%d chars, %d assets)", src.name, output.name, doc.stats.text_chars, assets.count)
    return Result(src, output, doc.stats, assets.count, use_pdf, findings)


_ID_LINE = re.compile(r"^id:\s*(.*?)\s*$", re.M)


_SOURCE_LINE = re.compile(r"^source:\s*(.*?)\s*$", re.M)


def _front_matter_field(output: Path, rx: re.Pattern[str]) -> str:
    if not output.is_file():
        return ""
    with output.open(encoding="utf-8", errors="replace") as f:
        head = f.read(4096)
    if not head.startswith("---\n"):
        return ""
    end = head.find("\n---\n", 4)
    m = rx.search(head[4:end] if end > 0 else "")
    if not m:
        return ""
    try:
        return json.loads(m.group(1)) if m.group(1).startswith('"') else m.group(1)
    except ValueError:
        return m.group(1)


def existing_id(output: Path) -> str:
    """前回の出力の Front Matter に人が付けた id があれば引き継ぐ（再変換で消さない）."""
    return _front_matter_field(output, _ID_LINE)


def existing_source(output: Path) -> str:
    """前回の出力の Front Matter にある元ファイル名（無ければ空）."""
    return _front_matter_field(output, _SOURCE_LINE)


def output_names(files: list[Path]) -> dict[Path, tuple[Path, str]]:
    """複数の入力を 1 つの出力先に書くときの、重ならない出力名を決める.

    戻り値は 入力 → (出力先からの相対フォルダ, 拡張子を除いた出力名)。入力の並び順によらず同じ結果になる。

    - 名前（拡張子を除く）が重ならなければ今までどおり `spec.md`
    - 拡張子だけが違う（spec.docx と spec.xlsx）なら拡張子を残す: `spec.docx.md` / `spec.xlsx.md`
    - 別フォルダの同じ名前（a/spec.docx と b/spec.docx）ならフォルダ構成を残す: `a/spec.md` / `b/spec.md`
    - それでも重なる（大文字小文字だけ違う等）なら `-2`, `-3` … を付ける

    Windows・macOS では大文字小文字を区別しないことが多いため、比較は大文字小文字を無視する。
    """
    files = sorted(set(files))
    groups: dict[str, list[Path]] = {}
    for f in files:
        groups.setdefault(f.stem.casefold(), []).append(f)

    names: dict[Path, tuple[Path, str]] = {}
    for group in groups.values():
        if len(group) == 1:
            names[group[0]] = (Path(), group[0].stem)
        elif len({f.name.casefold() for f in group}) == len(group):
            for f in group:
                names[f] = (Path(), f.name)
        else:
            parents = {f: f.parent.resolve() for f in group}
            common = Path(os.path.commonpath(list(parents.values())))
            by_sub: dict[Path, list[Path]] = {}
            for f in group:
                by_sub.setdefault(parents[f].relative_to(common), []).append(f)
            for sub, members in by_sub.items():
                for f in members:
                    names[f] = (sub, f.stem if len(members) == 1 else f.name)

    used: set[str] = set()
    for f in files:  # 最後の砦: 並び順で決まる連番
        sub, name = names[f]
        base, n = name, 1
        while (sub / f"{name}.md").as_posix().casefold() in used:
            n += 1
            name = f"{base}-{n}"
        used.add((sub / f"{name}.md").as_posix().casefold())
        names[f] = (sub, name)
    return names


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def front_matter(src: Path, title: str, doc_id: str = "") -> str:
    """AI や他ツールが文書を特定するための Front Matter.

    - id は人が付ける識別子（例: DOC-API-001）。変換では作らず、前回の出力にあれば引き継ぐ
    - source_sha256 は元ファイルの内容ハッシュ。元文書が更新されたか（ドリフト）の判定に使える
    - 日時は入れない（同じ入力から常に同じ出力にするため）
    """
    from . import __version__

    fields = {
        "title": title,
        "id": doc_id,
        "source": src.name,
        "source_sha256": sha256_of(src),
        "generator": f"doc2md {__version__}",
    }
    lines = [f"{k}: {json.dumps(v, ensure_ascii=False)}" for k, v in fields.items()]
    return "---\n" + "\n".join(lines) + "\n---\n\n"


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
    unique: dict[Path, Path] = {}  # 同じファイルを別の書き方で 2 回指定しても 1 回だけ変換する
    for f in sorted(set(files)):
        unique.setdefault(f.resolve(), f)
    return sorted(unique.values())


__all__ = [
    "Result",
    "convert_file",
    "collect_inputs",
    "output_names",
    "find_soffice",
    "office_to_pdf",
    "libreoffice_convert",
    "Options",
]
