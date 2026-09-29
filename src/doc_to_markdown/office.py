"""LibreOffice headless の呼び出し（PDF 化・Excel の再計算）."""

from __future__ import annotations

import subprocess
from pathlib import Path

from devtools_common.cli import UsageError
from devtools_common.executables import SOFFICE, find_executable, not_found_message


def find_soffice() -> str | None:
    return find_executable(SOFFICE)


def libreoffice_convert(src: Path, workdir: Path, target: str) -> Path:
    """src を target 形式に変換し、workdir 内の出力ファイルのパスを返す.

    workdir は呼び出し側が secure_tempdir で用意し、削除まで責任を持つ。
    target: "pdf" / "xlsx:Calc MS Excel 2007 XML" など（--convert-to に渡す値）
    """
    soffice = find_soffice()
    if not soffice:
        raise UsageError(not_found_message(SOFFICE))
    out_dir = workdir / "lo-out"
    out_dir.mkdir(exist_ok=True)
    profile = workdir / "lo-profile"  # 実行中の LibreOffice と衝突しないよう専用プロファイルを使う
    cmd = [
        soffice,
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--convert-to",
        target,
        "--outdir",
        str(out_dir),
        str(src),
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=600)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"LibreOffice の変換がタイムアウトしました: {src.name}") from e
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice の変換に失敗しました: {src.name} (exit {e.returncode})") from e
    ext = target.split(":", 1)[0]
    out = out_dir / f"{src.stem}.{ext}"
    if not out.is_file():
        raise RuntimeError(f"LibreOffice が {ext} を出力しませんでした: {src.name}")
    return out


def office_to_pdf(src: Path, workdir: Path) -> Path:
    return libreoffice_convert(src, workdir, "pdf")


def recalc_xlsx(src: Path, workdir: Path) -> Path:
    """数式の計算結果が保存されていないブックを、LibreOffice で開いて計算・保存し直す."""
    return libreoffice_convert(src, workdir, "xlsx:Calc MS Excel 2007 XML")
