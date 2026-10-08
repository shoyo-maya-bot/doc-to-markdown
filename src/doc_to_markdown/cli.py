"""doc2md — 各種ドキュメントを Markdown + 抽出図に変換する CLI.

例:
  doc2md 仕様書.docx -d out/
  doc2md docs/ -d out/ --recursive --format json      # バッチ + 機械可読レポート
  doc2md 見積.xlsx --header-row 4 -d out/            # 見出し行を指定（シート別: --header-row "4月=4"）
  doc2md scan.pdf --ocr-lang jpn -d out/             # スキャン PDF（要 Tesseract）
  doc2md 手順書.docx --via-pdf -d out/               # レイアウト優先（LibreOffice で PDF 経由）
  doc2md docs/ -d out/ --front-matter --llms-txt     # AI 向け: 識別情報付き + 全文書の索引/全文ファイル
  doc2md docs/ --dry-run                             # 変換せず、各ファイルをどう処理するかだけ表示

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

from . import __version__, llmstxt
from .adapters import OFFICE, SUPPORTED, VIA_PDF_ONLY
from .converter import collect_inputs, convert_file, output_names
from .model import Options

TOOL = "doc2md"
TESSERACT_TIMEOUT = 60  # --check-env で tesseract --list-langs を待つ秒数（固まったままにしない）


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
    p.add_argument(
        "--continue-on-error",
        action="store_true",
        help="1 ファイルの失敗（壊れたファイル・上限超過・外部ツールなし等）で止めずに続行する（バッチ用）。"
        "レポートと llms.txt は必ず出す。失敗したファイルがあれば終了コード 1（一部失敗）。引数の誤りは即終了コード 2",
    )
    p.add_argument(
        "--dry-run", action="store_true", help="変換せず、各ファイルの処理方法（読み取り方式・必要な外部ツール）を表示"
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    a = p.add_argument_group("AI 向けの出力")
    a.add_argument(
        "--front-matter",
        action="store_true",
        help="Markdown の先頭に title / id / source / source_sha256 を付ける（id は前回の出力から引き継ぐ）",
    )
    a.add_argument(
        "--llms-txt",
        action="store_true",
        help="出力先に llms.txt（文書一覧）と llms-full.txt（全文の連結）も作る",
    )

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
    g.add_argument(
        "--max-unzipped-mb",
        type=int,
        default=1024,
        help="Office 形式（zip）の展開後の合計サイズの上限 MB。超えたら変換しない（終了コード 2。既定 1024）",
    )
    g.add_argument(
        "--max-zip-ratio",
        type=int,
        default=500,
        help="Office 形式の中の 1MB 以上のファイルの圧縮率の上限。超えたら変換しない（終了コード 2。既定 500）",
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
        max_unzipped_mb=args.max_unzipped_mb,
        max_zip_ratio=args.max_zip_ratio,
        ocr=args.ocr,
        ocr_lang=args.ocr_lang,
        remove_running_headers=not args.keep_running_headers,
        front_matter=args.front_matter,
    )


def plan(files: list[Path], opts: Options) -> list[Finding]:
    """--dry-run: 各ファイルをどの方式で読むか、外部ツールが要るかを示す（ファイルは書かない）."""
    from devtools_common.executables import SOFFICE, find_executable

    have_soffice = find_executable(SOFFICE) is not None
    out: list[Finding] = []
    for f in files:
        suffix = f.suffix.lower()
        if suffix not in SUPPORTED:
            out.append(Finding("plan.unsupported", Severity.WARNING, "未対応の形式のため変換しません", location=f.name))
            continue
        via_pdf = suffix in VIA_PDF_ONLY or (opts.via_pdf and suffix in OFFICE)
        route = f"{suffix} → PDF 経由（LibreOffice）" if via_pdf else f"{suffix} を直接読み取り"
        if via_pdf and not have_soffice:
            out.append(
                Finding(
                    "plan.missing_tool",
                    Severity.ERROR,
                    f"{route}: LibreOffice が見つかりません",
                    location=f.name,
                    suggestion="LibreOffice を入れるか DEVTOOLS_SOFFICE で場所を指定（doc2md --check-env）",
                )
            )
        else:
            out.append(Finding("plan.route", Severity.INFO, route, location=f.name))
    return out


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

        try:
            out = subprocess.run(
                [tess, "--list-langs"], capture_output=True, text=True, timeout=TESSERACT_TIMEOUT
            ).stdout
        except (subprocess.TimeoutExpired, OSError) as e:
            print(f"       OCR 言語: 取得できませんでした（{type(e).__name__}）")
            return common_cli.ExitCode.OK
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
    files = collect_inputs(args.inputs, recursive=args.recursive, out_dir=args.out_dir)
    if not files:
        raise common_cli.UsageError("変換対象のファイルがありません")

    if args.dry_run:
        findings = plan(files, opts)
        text = report.render(
            findings,
            args.format,
            title="doc2md 変換計画（dry-run）",
            meta={"inputs": len(files)},
            tool=TOOL,
            tool_version=__version__,
        )
        common_cli.write_output(text, args.output)
        return common_cli.ExitCode.FINDINGS if report.has_errors(findings) else common_cli.ExitCode.OK

    # 使い方の誤り（未対応の形式を直接指定した等）は、1 ファイルも書く前に止める（--continue-on-error でも続行しない）
    unsupported = [f.name for f in files if f.suffix.lower() not in SUPPORTED]
    if unsupported:
        raise common_cli.UsageError(
            f"未対応の形式です: {', '.join(unsupported)}（対応: {', '.join(sorted(SUPPORTED))}）"
        )

    findings: list[Finding] = []
    converted = []
    input_errors = 0  # --continue-on-error で続けた、ファイル単位の入力の誤り（UsageError）の数
    names = output_names(files)
    for f in files:
        sub, name = names[f]
        if (sub, name) != (Path(), f.stem):  # 同じ出力名の入力があるので名前を変える（黙って上書きしない）
            dest = (sub / f"{name}.md").as_posix()
            print(f"doc2md: 出力名が他の入力と重なるため {f.as_posix()} → {dest} に書き出します", file=sys.stderr)
            findings.append(
                Finding(
                    "convert.renamed",
                    Severity.INFO,
                    f"出力名が他の入力と重なるため {dest} に書き出します",
                    location=f.name,
                )
            )
        try:
            r = convert_file(f, args.out_dir / sub, options=opts, name=name)
        except common_cli.UsageError as e:
            # そのファイルだけの入力の誤り（zip の上限超過、LibreOffice が無い .doc 等）.
            # --continue-on-error なら、そのファイルの失敗として記録して続ける（終了コードは最後に 1 = 一部失敗）
            if not args.continue_on_error:
                raise
            input_errors += 1
            findings.append(
                Finding("convert.input_error", Severity.ERROR, str(e), location=f.name, suggestion="入力ファイルを確認")
            )
            continue
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

    # --continue-on-error で失敗したファイルがあっても、変換できた分の索引は必ず作る（0 件なら空の索引）
    if args.llms_txt:
        written = llmstxt.write(args.out_dir, [(r.source.name, r.output) for r in converted])
        findings.append(Finding("convert.llms_txt", Severity.INFO, "→ " + ", ".join(w.as_posix() for w in written)))

    text = report.render(
        findings,
        args.format,
        title="doc2md 変換レポート",
        meta={"inputs": len(files), "converted": len(converted), "out_dir": args.out_dir.as_posix()},
        tool=TOOL,
        tool_version=__version__,
    )
    common_cli.write_output(text, args.output)

    # 続行したファイル単位の入力の誤り（convert.input_error）は一部失敗として終了コード 1.
    # 2 は引数の誤りなど、何も変換せずに止めた場合だけに使う
    if input_errors or report.has_errors(findings):
        return common_cli.ExitCode.FINDINGS
    if args.strict and any(f.severity is Severity.WARNING for f in findings):
        return common_cli.ExitCode.FINDINGS
    return common_cli.ExitCode.OK


def entrypoint() -> None:
    sys.exit(common_cli.run_main(main))


if __name__ == "__main__":
    entrypoint()
