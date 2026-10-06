# Changelog

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)、版は [SemVer](https://semver.org/lang/ja/) に従います。

## [Unreleased]

### Fixed

- Excel: セルの表示形式を無視していた問題を修正。`000` の表示形式で 1 と入ったコードが `1` になる、
  日付がいつも `2026-04-01` 形式になる、などが起きていた
  - ゼロ埋め（`000`）・小数桁（`0.00`）・桁区切り（`#,##0`）・パーセント・通貨記号・文字列（`@`）・
    日付時刻（`yyyy/mm/dd`、`m/d`、`yyyy"年"m"月"d"日"`、`h:mm` 等）・和暦（`ggge"年"`）を表示形式どおりに出す
  - **出力が変わります**: 例えば表示形式 `yyyy/mm/dd` の日付は `2026-04-01` ではなく `2026/04/01` になる。
    表示形式が「標準」のセルと、地域設定に従う既定の日付は従来どおり
  - 再現できない表示形式（指数・分数・条件付き 等）は従来どおりの値を出し、`loss.number_format`（warning）で知らせる

## [0.4.1] - 2026-10-05

### Fixed

- 出力名が重なる複数の入力（`spec.docx` と `spec.xlsx`、別フォルダの同じ名前）を 1 回で変換すると、
  後のファイルが前の出力と抽出図を黙って上書きしていた問題を修正
  - 拡張子だけ違うときは `spec.docx.md` / `spec.xlsx.md`、別フォルダの同じ名前は `a/spec.md` / `b/spec.md` に書き出す
  - 名前を変えたときは標準エラーとレポート（`convert.renamed`）で知らせる。重ならない入力の名前は変わらない
- `--front-matter`: 既存の出力が別の元ファイルから作られていたとき、その `id` を引き継いでいた問題を修正。
  上書きしたことを `convert.overwrite_other_source`（warning）で知らせる
- 同じファイルを別の書き方（`docs/a.docx` と `./docs/a.docx`）で 2 回指定すると 2 回変換していた問題を修正

## [0.4.0] - 2026-10-02

### Added

- `figspec`: 図の定義（YAML / JSON）から編集できる draw.io を生成。線のルール検査（宙に浮いた線・自己ループ・重複・判断のラベル・
  孤立ノード・重なり）、自動配置、`--svg`。Copilot の `/figspec-draw`
- docs/MIGRATION_GUIDE.md: 設計書を Markdown に移すロット運用と突合チェックリスト

## [0.3.0] - 2026-10-02

### Added

- `--front-matter`: 出力 Markdown に title / id / source / source_sha256 / generator を付与（id は再変換で引き継ぐ）
- `--llms-txt`: 出力先に llms.txt（文書一覧）と llms-full.txt（全文連結）を作成
- `--dry-run`: 変換せず、各ファイルの読み取り方式と不足している外部ソフトを表示
- `scripts/verify_gates.py`、Copilot への指示に AI エージェントの約束

## [0.2.0] - 2026-09-30

### Added

- Word: ハイパーリンク、テキストボックス、脚注・文末脚注、コメント、変更履歴（`--track-changes`）、
  コンテンツコントロール、縦結合・入れ子の表、ヘッダー・フッター（`--no-header-footer`）、数式の文字
- Excel: タイトル行・注記行の分離、見出し行の自動検出と `--header-row`、2 段見出し、セル結合、1 シート複数表、
  非表示の行・列・シート（`--skip-hidden`）、数式の計算結果が無いブックの検知と `--recalc`、パーセント表示、
  セル数による省メモリ読み込み（`--max-cells`）
- PowerPoint: SmartArt の文字
- PDF: スキャンページの OCR（Tesseract。`--ocr` / `--ocr-lang`）、2 段組み、繰り返しヘッダー・フッター・ページ番号の除去、
  太字・「第N条」からの見出し推定、ページ単位のメモリ解放
- 元ファイルから直接数えた文字数による変換ロス検知（`loss.raw_text`）、`convert.warning`
- `--check-env`、setup.cmd / setup.sh、Copilot のプロンプト（`/doc2md-convert` 他）、docs/OUTPUT_SPEC.md、scripts/bench.py
- 実務に近いテスト文書（samples/realworld、LibreOffice で保存）とゴールデンテスト

### Changed

- アダプタの引数に `Options` を追加（`convert(path, assets, opts)`）。ブロックは逐次生成・逐次書き出し
- 外部ソフトの探索を devtools-common の `executables` に統一（Windows の Program Files 等も探す）
- レポートの JSON に `schema_version` / `tool` / `tool_version`（devtools-common 0.2.0）
- 表のセル内の改行を `<br>` で保持

## [0.1.0] - 2026-09-28

### Added

- 初版
