"""doc2md — 各種ドキュメントを Markdown + 抽出図に変換する CLI.

例:
  doc2md 仕様書.docx -d out/
  doc2md docs/ -d out/ --recursive --format json      # バッチ + 機械可読レポート
  doc2md 見積.xlsx --via-pdf -d out/                  # レイアウト優先（LibreOffice で PDF 経由）

標準出力には変換レポート（ロス検知・注意事項）を出す。Markdown はファイルに書き出す。
終了コード: 0 成功 / 1 変換ロスの疑い（--strict 時）/ 2 引数・入力の誤り / 3 実行時エラー
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from devtools_common import cli as common_cli
from devtools_common import report
from devtools_common.report import Finding, Severity

from . import __version__
from .adapters import SUPPORTED
from .converter import collect_inputs, convert_file


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="doc2md", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    p.add_argument(
        "inputs", nargs="+", type=Path, help=f"変換するファイルまたはディレクトリ（{' '.join(sorted(SUPPORTED))}）"
    )
    p.add_argument("-d", "--out-dir", type=Path, default=Path("out"), help="Markdown と抽出図の出力先（既定: out/）")
    p.add_argument("-r", "--recursive", action="store_true", help="ディレクトリを再帰的に探索する")
    p.add_argument("--via-pdf", action="store_true", help="Office 文書を LibreOffice で PDF 化してから抽出する")
    p.add_argument("--strict", action="store_true", help="変換ロスの疑い（warning）があれば終了コード 1")
    p.add_argument("--continue-on-error", action="store_true", help="1 ファイルの失敗で止めずに続行する（バッチ用）")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    common_cli.add_common_args(p)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    common_cli.configure_from_args(args)
    files = collect_inputs(args.inputs, recursive=args.recursive)
    if not files:
        raise common_cli.UsageError("変換対象のファイルがありません")

    findings: list[Finding] = []
    converted = []
    for f in files:
        try:
            r = convert_file(f, args.out_dir, via_pdf=args.via_pdf)
        except common_cli.UsageError:
            raise
        except Exception as e:  # noqa: BLE001
            if not args.continue_on_error:
                raise
            findings.append(Finding("convert.failed", Severity.ERROR, f"{type(e).__name__}: {e}", location=f.name))
            continue
        converted.append(r)
        findings.append(
            Finding(
                "convert.ok",
                Severity.INFO,
                f"→ {r.output.as_posix()}",
                location=f.name,
                actual=f"headings={r.stats.headings} lists={r.stats.list_items} tables={r.stats.tables} "
                f"rows={r.stats.table_rows} images={r.stats.images} chars={r.stats.text_chars}"
                + (" via-pdf" if r.via_pdf else ""),
            )
        )
        findings.extend(r.findings)

    text = report.render(
        findings,
        args.format,
        title="doc2md 変換レポート",
        meta={"inputs": len(files), "converted": len(converted), "out_dir": args.out_dir.as_posix()},
    )
    if args.output == "-":
        sys.stdout.write(text)
    else:
        Path(args.output).write_text(text, encoding="utf-8", newline="\n")

    if report.has_errors(findings):
        return common_cli.ExitCode.FINDINGS
    if args.strict and any(f.severity is Severity.WARNING for f in findings):
        return common_cli.ExitCode.FINDINGS
    return common_cli.ExitCode.OK


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))


if __name__ == "__main__":
    entrypoint()
