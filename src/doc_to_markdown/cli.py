"""doc2md — 各種ドキュメントを Markdown + 抽出図に変換する CLI.

例:
  doc2md 仕様書.docx -d out/
  doc2md docs/ -d out/ --recursive --format json      # バッチ + 機械可読レポート
  doc2md 見積.xlsx --header-row 4 -d out/            # 見出し行を指定（シート別: --header-row "4月=4"）
  doc2md scan.pdf --ocr-lang jpn -d out/             # スキャン PDF（要 Tesseract）
  doc2md 手順書.docx --via-pdf -d out/               # レイアウト優先（LibreOffice で PDF 経由）

標準出力には変換レポート（ロス検知・注意事項）を出す。Markdown はファイルに書き出す。
終了コード: 0 成功 / 1 失敗あり（--strict 時は警告ありも）/ 2 引数・入力の誤り / 3 実行時エラー
外部ツールの場所: DEVTOOLS_SOFFICE / DEVTOOLS_DRAWIO / DEVTOOLS_TESSERACT（PATH に無い場合）
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
from .model import Options

TOOL = "doc2md"


def _header_rows(values: list[str]) -> dict[str, int]:
    out: dict[str, int] = {}
    for v in values:
        sheet, _, row = v.rpartition("=")
        try:
            n = int(row)
        except ValueError:
            raise common_cli.UsageError(f"--header-row の値が不正です: {v}（例: 4 または 4月=4）") from None
        if n < 1:
            raise common_cli.UsageError(f"--header-row は 1 以上: {v}")
        out[sheet or "*"] = n
    return out


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog=TOOL, description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("inputs", nargs="*", type=Path, help=f"ファイルまたはディレクトリ（{' '.join(sorted(SUPPORTED))}）")
    p.add_argument(
        "--check-env",
        action="store_true",
        help="任意の外部ツール（LibreOffice・Tesseract・draw.io）の有無を表示して終了",
    )
    p.add_argument("-d", "--out-dir", type=Path, default=Path("out"), help="Markdown と抽出図の出力先（既定: out/）")
    p.add_argument("-r", "--recursive", action="store_true", help="ディレクトリを再帰的に探索する")
    p.add_argument("--strict", action="store_true", help="警告（変換ロスの疑い等）があれば終了コード 1")
    p.add_argument("--continue-on-error", action="store_true", help="1 ファイルの失敗で止めずに続行する（バッチ用）")
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    g = p.add_argument_group("Office 共通")
    g.add_argument("--via-pdf", action="store_true", help="LibreOffice で PDF 化してから抽出する（レイアウト優先）")

    w = p.add_argument_group("Word")
    w.add_argument(
        "--track-changes",
        choices=["accept", "reject", "show"],
        default="accept",
        help="変更履歴: accept=承認後（既定）/ reject=変更前 / show=差分を ~~削除~~ <ins>挿入</ins> で表示",
    )
    w.add_argument("--no-comments", action="store_true", help="コメントを出力しない")
    w.add_argument("--no-header-footer", action="store_true", help="ヘッダー・フッターを出力しない")

    x = p.add_argument_group("Excel")
    x.add_argument(
        "--header-row",
        action="append",
        default=[],
        metavar="[シート=]行",
        help="見出し行（1 始まり）。既定は自動検出。シート別指定は 4月=4（複数指定可）",
    )
    x.add_argument("--skip-hidden", action="store_true", help="非表示の行・列・シートを出力しない")
    x.add_argument(
        "--recalc", action="store_true", help="数式の計算結果が無いブックを LibreOffice で再計算してから読む"
    )
    x.add_argument(
        "--max-cells",
        type=int,
        default=200_000,
        help="これを超えるセル数のブックは省メモリ読み込み（結合・複数表の判定なし。既定 200000）",
    )

    d = p.add_argument_group("PDF")
    d.add_argument(
        "--ocr",
        choices=["auto", "always", "never"],
        default="auto",
        help="OCR: auto=文字の無いページだけ（既定）/ always=全ページ / never=しない（要 Tesseract）",
    )
    d.add_argument("--ocr-lang", default="jpn+eng", help="OCR の言語（Tesseract の言語コード。既定 jpn+eng）")
    d.add_argument("--keep-running-headers", action="store_true", help="全ページに繰り返し出る行を除去しない")
    common_cli.add_common_args(p)
    return p


def options_from_args(args: argparse.Namespace) -> Options:
    return Options(
        via_pdf=args.via_pdf,
        track_changes=args.track_changes,
        include_comments=not args.no_comments,
        include_header_footer=not args.no_header_footer,
        header_rows=_header_rows(args.header_row),
        skip_hidden=args.skip_hidden,
        recalc=args.recalc,
        xlsx_full_load_max_cells=args.max_cells,
        ocr=args.ocr,
        ocr_lang=args.ocr_lang,
        remove_running_headers=not args.keep_running_headers,
    )


def check_env() -> int:
    """外部ツールの有無と、使えない機能を表示する（どれも無くても基本の変換は動く）."""
    from devtools_common.executables import DRAWIO, SOFFICE, TESSERACT, find_executable, not_found_message

    rows = [
        (SOFFICE, "--via-pdf、.doc/.xls/.ppt 等の旧形式、--recalc"),
        (TESSERACT, "スキャン PDF の OCR"),
        (DRAWIO, "draw.io 図の svg 書き出し（無くても図の文字は抽出）"),
    ]
    print(f"doc2md {__version__} / Python {sys.version.split()[0]}")
    for spec, feature in rows:
        path = find_executable(spec)
        mark = "OK  " if path else "なし"
        print(f"[{mark}] {spec.key:10} {path or not_found_message(spec)}")
        print(f"       用途: {feature}")
    tess = find_executable(TESSERACT)
    if tess:
        import subprocess

        out = subprocess.run([tess, "--list-langs"], capture_output=True, text=True).stdout
        langs = [ln.strip() for ln in out.splitlines()[1:] if ln.strip()]
        print(
            f"       OCR 言語: {' '.join(langs) or '(なし)'}"
            + ("" if "jpn" in langs else "  ← 日本語(jpn)がありません")
        )
    return common_cli.ExitCode.OK


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    common_cli.configure_from_args(args)
    if args.check_env:
        return check_env()
    if not args.inputs:
        raise common_cli.UsageError("変換するファイルかディレクトリを指定してください（doc2md --help）")
    opts = options_from_args(args)
    files = collect_inputs(args.inputs, recursive=args.recursive)
    if not files:
        raise common_cli.UsageError("変換対象のファイルがありません")

    findings: list[Finding] = []
    converted = []
    for f in files:
        try:
            r = convert_file(f, args.out_dir, options=opts)
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
        tool=TOOL,
        tool_version=__version__,
    )
    common_cli.write_output(text, args.output)

    if report.has_errors(findings):
        return common_cli.ExitCode.FINDINGS
    if args.strict and any(f.severity is Severity.WARNING for f in findings):
        return common_cli.ExitCode.FINDINGS
    return common_cli.ExitCode.OK


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))


if __name__ == "__main__":
    entrypoint()
