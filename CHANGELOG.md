# Changelog

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)、版は [SemVer](https://semver.org/lang/ja/) に従います。

## [Unreleased]

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
