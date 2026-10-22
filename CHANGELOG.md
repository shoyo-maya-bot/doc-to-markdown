# Changelog

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/)、版は [SemVer](https://semver.org/lang/ja/) に従います。

## [Unreleased]

### Changed（破壊的変更）

- **Python 3.10 のサポートをやめ、3.11 以上が必要になった**（`requires-python = ">=3.11"`）。
  Pillow 12.3.0 の脆弱性（2026-10-21 公開、CRITICAL を含む 8 件）の修正版 13.0.0 が Python 3.11 以上でしか入らないため。
  3.10 の環境では `pip install` が失敗する（`setup.cmd`・`setup.sh` も 3.11 未満で止まる）ので、Python 3.11 以上で `.venv` を作り直す。
  CI の Linux のテストは 3.11 と 3.12、constraints.txt の基準も 3.11 にした

### Security

- pillow を 12.3.0 から 13.0.0 に上げる（constraints.txt）

### Fixed（第 7 回）

- Excel・PowerPoint・PDF・draw.io の文字（素の文字）で、行の途中の `*` `_` `` ` `` `[` `]` `~` をエスケープしていなかった
  （`2*3*4` が斜体に、`[x](javascript:…)` がリンクになった）。前に `\` を付ける（表示される文字は同じ。Word は従来どおり）。
  **Markdown のソースが変わります**（例: `a_b` → `a\_b`）。fieldcheck の表の読み取りは `\` を外して元の文字に戻す
- ロス検知で、エスケープした `\[x\](…)` をリンクとみなして文字を落とし、`loss.text` / `loss.raw_text` を誤って出していた
- `-r` で別フォルダに同じ名前のファイルがあると、レポートの `location` がファイル名だけで区別できなかった。
  重なるときだけフォルダ付き（`a/spec.docx`）にする（重ならない名前は従来どおり。`--dry-run` も同じ）
- `--max-cells` / `--max-unzipped-mb` / `--max-zip-ratio` に 0 以下を渡せた（0 ですべての Office 文書が上限超えになる等）。
  1 以上でなければ引数の誤り（終了コード 2）にする

### Changed（第 7 回: ドキュメントを実装に合わせる）

- 版を `0.5.0.dev0` にした（Unreleased に破壊的変更（3.11 以上）があるのに、リリース済みの `0.4.1` を名乗っていた）
- OUTPUT_SPEC.md: `convert.llms_txt` が rule_id の一覧に無かった。終了コード 2 と `convert.input_error` に壊れたファイルを書いた。
  素の文字のインライン記法のエスケープと、重なる名前の `location` を書いた
- WORKFLOW.md・copilot-instructions.md: Tesseract が無いと終了コード 2 になると書いていた（実際は止まらず `convert.warning`）
- README の上限の表に「1 以上」を書いた
- `tests/test_docs_match_code.py`: Python の版（pyproject・README・setup・CI・ruff）、版と CHANGELOG、README の上限の既定値と
  `--help` の既定値、OUTPUT_SPEC の rule_id とコード、終了コードの一致を確かめる

### Changed（第 6 回）

- devtools-common の固定を `04f3915` に更新（設定ファイルの構文エラーや予期しないエラーの標準エラー出力に伏せ字を通す等）

### Fixed（第 6 回）

- PDF のヘッダー・フッター除去で、数字だけが違う本文の行（ページの下端の `Total 1200 yen` など）が消えていたのを直す。
  数字違いの行は「全ページに出て、数字がページ番号と同じ歩みで増える」ときだけ除き、同じ文字の行は半数以上のページに出たときに除く。
  除いた文字数はロス検知（`loss.raw_text`）の元の文字数から引かない（ページが少ない PDF では warning が出ることがある）
- `llms.txt` のリンクの文字で `[` `]` をエスケープしていなかった（`a]b(c` の見出しでリンクが壊れた）。元ファイル名もエスケープし、
  `llms-full.txt` のコメントの `--` を崩す
- `--via-pdf`・`--recalc` で `-` で始まるファイル名が LibreOffice のオプションとして読まれないよう、絶対パスで渡す
- 仕様書（OUTPUT_SPEC.md）のヘッダー・フッター除去の条件をコードに合わせる

### Added（第 5 回）

- `--loss-threshold PERCENT`（既定 95）: 本文の文字数のロス検知（`loss.text` / `loss.raw_text`）のしきい値を変えられる。
  文字数だけを見るため小さな欠落・入れ替わり・文字化けは検知できないことを README・WORKFLOW に明記

### Fixed（第 5 回）

- Word の番号付き箇条書きが段落や表を挟むたびに 1 に戻っていたのを、Word と同じ続き番号にする
  （numId・段ごと。同じ抽象番号の一覧は続け、`startOverride` で振り直し、`w:start` を反映）
- 大きなブックの省メモリ読み込みで `--header-row` と `--skip-hidden` が黙って無視されていたのを、どちらも効くようにする
  （非表示の行・列はシートの XML から読む。判定できなければ `convert.hidden_unknown`（warning））。
  省メモリ読み込みでも、非表示の行を含めて出力したときは `convert.warning` で知らせる
- 前回の出力が BOM 付きで保存し直されていると Front Matter の `id`・`source` を読めず、`id` を失っていたのを直す
  （utf-8-sig で読む。llms.txt の作成も同様）

### Fixed（第 4 回）

- 壊れた入力（拡張子だけ .docx のテキスト、途中で切れた zip・PDF、Word の部品の無い zip、壊れた draw.io の XML 等）を
  入力の誤り（終了コード 2）として扱う（これまでは実行時エラー 3）。`--continue-on-error` では一部失敗（1）
- 脚注・文末脚注・コメントの本文も、本文と同じく `<`・`<!--`・文字参照をエスケープする（OUTPUT_SPEC どおり）
- 出力先が入力フォルダの中にあるとき、フォルダの探索で出力先（とその `.assets/`）の中を拾わない
  （再実行で前回の抽出図や draw.io のソースの写しを変換していた）
- draw.io の圧縮された図の展開上限を、ページごとだけでなくファイル全体の合計でも見る
- `--check-env` の `tesseract --list-langs` に時間制限を付け、OCR の時間切れはそのページを読まずに警告する

### Changed（第 4 回）

- devtools-common の固定を `6a263b8` に更新

### Changed

- `--continue-on-error` で一部のファイルが入力の誤り（`convert.input_error`）で失敗したときの終了コードを 2 → 1（一部失敗）に変更。
  全ファイル成功なら 0、引数の誤り・未対応の形式の直接指定は従来どおり即終了コード 2

### Fixed（第 3 回）

- 画像・draw.io のソース・ハイパーリンク・llms.txt のリンク先に空白・括弧を含むパスや URL を入れても切れないよう、
  その文字だけをパーセントエンコードする（`図 (1).assets/…` → `図%20%281%29.assets/…`。日本語はそのまま）
- Word（markup=True の段落・見出し・箇条書き・表・脚注・コメント）でも、本文の `\` と Markdown 記号（`` ` `` `*` `_` `[` `]` `~`）を
  エスケープする（太字・リンク・変更履歴など変換側の記法は残す）。本文の文字数（ロス検知）にはエスケープの `\` を数えない
- `--continue-on-error` で、1 ファイルの入力の誤り（zip の上限超え、LibreOffice の無い .doc 等）で全体が止まっていた →
  `convert.input_error`（error）として続行し、レポートと llms.txt を必ず出す（終了コードは 2。引数の誤りは従来どおり即終了）
- draw.io の圧縮された図（deflate）の展開後のサイズに上限（`--max-unzipped-mb` と同じ値）を設ける（圧縮爆弾対策）

### Changed（第 3 回）

- devtools-common の固定を `cbe97d5` に更新
- CI にカバレッジの測定と下限（85%。pyproject.toml の `[tool.coverage.report]`）を追加

### Changed

- Excel の読み取りの共通部分（表示形式の反映 `render_value`、結合セルの索引、複数段の見出しの判定）を
  devtools-common の `devtools_common.excel` に移した（design-diff と同じ規則で読む）。
  `adapters/xlsx_numfmt.py` は削除し、そのテストも devtools-common へ移した。devtools-common は `[excel]` extra 付きで固定。
  出力は変わらない（ゴールデンファイルは同一）。見出しの段数の判定は、縦結合の下の行が数値だけのときは子見出しとみなさない
  （3 段に縦結合した見出しで、3 段目がデータの表を 3 段の見出しと誤認していた）

### Added

- `--max-unzipped-mb`（既定 1024）と `--max-zip-ratio`（既定 500）: Office 形式（zip）の展開後の合計サイズと圧縮率の上限。
  超えたら変換せずに終了コード 2 で止める（zip 爆弾対策）
- `loss.table_as_paragraph`（warning）: Excel の表らしい範囲で見出し行が見つからず段落にしたことを知らせる
- `convert.large_workbook`（warning）: 大きなブックを省メモリで読んだことを知らせる（これまでは info の `convert.note`）

### Fixed

- `--continue-on-error` でも、使い方の誤り（未対応の形式を直接指定・展開サイズの上限超過など。終了コード 2）は続行せずに止める。
  未対応の形式は 1 ファイルも書き出す前に止める
- 箇条書き・引用（テキストボックス・SmartArt）・HTML コメントの文字も Markdown としてエスケープする
  （段落・見出し・表と同じ。`# 〜` が見出しに、`<!-- … -->` が描画で消えていた）
- Excel: 表の末尾の注記らしい行（※・注 等で始まる 1 セルの行）を表から外したときに警告を出す（外す条件は従来どおり）
- 再変換で失敗したとき、前回の抽出図（`<name>.assets/`）を消さずに残す（古い抽出図は変換が成功してから消す）
- Excel: 非表示の行の数え方が行数 × セル数に比例していた（大きなブックで遅い）のを、値のある行の集合で数えるように

- Excel: 見出しが `項目 / 2024 / 2025 / 2026` のような年度の表が、表として認識されず段落の羅列になっていた問題を修正。
  年度の並び・太字・文字列として入力された数値で見出し行を見分ける
- Excel: 1 列だけの表が段落になっていた問題を修正（見出しが太字、または本文がすべて数値で 3 行以上なら表）
- Excel: 経過時間の表示形式（`[h]:mm` 等）が `1 day, 12:00:00` になっていた問題を修正（`36:00` と出す）
- Excel: 省メモリ読み込み（`--max-cells` 超え）で表示形式を反映していなかった（`001` が `1` になる）問題を修正
- 段落・見出しの文字が Markdown の記法として解釈されていた問題を修正。セル値 `# 大見出し` が見出しに、
  `- 〜` `1. 〜` が箇条書きに、`<!-- … -->` が描画で消えていた。行頭の記法と HTML をエスケープする
  - **出力が変わります**: 該当する文字のある段落・見出し・表のセルに `\` が付く（docs/OUTPUT_SPEC.md の「エスケープ」）
- 変換の途中で失敗すると、途中までの `.md` が残っていた問題を修正（一時ファイルに書いてから置き換える）
- Excel: 大きなブックの省メモリ読み込みを info でしか知らせていなかった（`--strict` で検知できなかった）

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
