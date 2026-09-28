"""変換の流れ: (前処理) → アダプタ → 正規化 → 描画 → ロス検知 → 書き出し."""

from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

from devtools_common.cli import UsageError
from devtools_common.log import get_logger
from devtools_common.report import Finding, Severity
from devtools_common.tempfiles import secure_tempdir

from . import lossaudit
from .adapters import ADAPTERS, OFFICE, SUPPORTED, VIA_PDF_ONLY, pdf_adapter
from .assets import AssetWriter
from .model import Document, Heading, SourceStats
from .normalize import normalize
from .render import render

log = get_logger(__name__)

SOFFICE_BINARIES = ("soffice", "libreoffice")


@dataclass
class Result:
    source: Path
    output: Path
    stats: SourceStats
    assets: int
    via_pdf: bool
    findings: list[Finding] = field(default_factory=list)


def find_soffice() -> str | None:
    return next((p for p in (shutil.which(b) for b in SOFFICE_BINARIES) if p), None)


def office_to_pdf(src: Path, workdir: Path) -> Path:
    """LibreOffice headless で PDF 化する（workdir は呼び出し側が安全に削除する一時ディレクトリ）."""
    soffice = find_soffice()
    if not soffice:
        raise UsageError("LibreOffice (soffice) が見つかりません。--via-pdf / 旧形式の変換にはインストールが必要です")
    profile = workdir / "lo-profile"  # 実行中の LibreOffice と衝突しないよう専用プロファイルを使う
    cmd = [
        soffice,
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(workdir),
        str(src),
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=300)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"LibreOffice の変換がタイムアウトしました: {src.name}") from e
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice の変換に失敗しました: {src.name} (exit {e.returncode})") from e
    pdf = workdir / f"{src.stem}.pdf"
    if not pdf.is_file():
        raise RuntimeError(f"LibreOffice が PDF を出力しませんでした: {src.name}")
    return pdf


def _audit_markdown(doc: Document) -> str:
    return render(normalize([b for b in doc.blocks if not getattr(b, "synthetic", False)]))


def convert_file(src: Path, out_dir: Path, *, via_pdf: bool = False) -> Result:
    suffix = src.suffix.lower()
    if suffix not in SUPPORTED:
        raise UsageError(f"未対応の形式です: {src.name}（対応: {', '.join(sorted(SUPPORTED))}）")
    use_pdf = suffix in VIA_PDF_ONLY or (via_pdf and suffix in OFFICE)
    stem = src.stem
    out_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = out_dir / f"{stem}.assets"
    if assets_dir.exists():  # 再実行時に古い抽出図が残らないようにする（決定的な出力）
        shutil.rmtree(assets_dir)
    assets = AssetWriter(out_dir, stem)

    if use_pdf:
        with secure_tempdir("doc2md-") as tmp:
            pdf = office_to_pdf(src, tmp)
            doc = pdf_adapter.convert(pdf, assets)
        # PDF のファイル名由来の H1 を元ファイル名に戻す
        if doc.blocks and isinstance(doc.blocks[0], Heading) and doc.blocks[0].synthetic:
            doc.blocks[0] = Heading(1, stem, synthetic=True)
    else:
        doc = ADAPTERS[suffix](src, assets)

    md = render(normalize(doc.blocks))
    output = out_dir / f"{stem}.md"
    output.write_text(md, encoding="utf-8", newline="\n")

    findings = lossaudit.audit(doc.stats, _audit_markdown(doc), location=src.name)
    findings += [Finding("convert.note", Severity.INFO, n, location=src.name) for n in doc.notes]
    log.info("converted %s -> %s (%d chars, %d assets)", src.name, output.name, doc.stats.text_chars, assets.count)
    return Result(src, output, doc.stats, assets.count, use_pdf, findings)


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


__all__ = ["Result", "convert_file", "collect_inputs", "office_to_pdf"]
