# doc-to-markdown（doc2md）

各種ドキュメント（Office / PDF / 図）を **Markdown + 抽出図** に変換することだけに責務を絞ったツールです。
内容の検証（項目間の整合性チェック）は行いません → [fieldcheck](https://github.com/shoyo-maya-bot/office-review/tree/main/fieldcheck)（[office-review](https://github.com/shoyo-maya-bot/office-review) に同梱）に別プロセスで渡します。

用途の例: 仕様書・手順書を Git 管理下のテキストにする／生成 AI（Copilot）に読ませる前処理／差分レビュー／検索

## すぐ使う（3 ステップ）

1. このリポジトリを clone する（`devtools-common` も同じフォルダに clone しておくと GitHub の認証が不要）
2. **Windows**: `setup.cmd` をダブルクリック ／ **macOS・Linux**: `./setup.sh`
3. VS Code でこのフォルダを開き、Copilot Chat で **`/doc2md-convert`** と入力 → ファイルのパスを指定

| Copilot のプロンプト | 何をするか |
|---|---|
| `/doc2md-convert` | 変換して、注意点（変更履歴・非表示行・OCR など）と中身の概要を報告 |
| `/doc2md-investigate-loss` | 「内容が足りない」ときに、通常の変換と PDF 経由の変換を比べて原因を調べる |

ターミナルから直接使う場合は、`.venv` を有効にしてから `doc2md --help`。

### あると便利な無料ソフト（無くても基本の変換は動く）

| ソフト | 使う機能 | 確認 |
|---|---|---|
| [LibreOffice](https://www.libreoffice.org/download/) | `--via-pdf`、旧形式（.doc/.xls/.ppt）、`--recalc` | `doc2md --check-env` |
| [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)（日本語データ付き） | スキャン PDF の OCR | 同上 |
| [draw.io Desktop](https://github.com/jgraph/drawio-desktop/releases) | 図の svg 書き出し | 同上 |

標準の場所（Windows の Program Files 等）以外に入れた場合は、環境変数 `DEVTOOLS_SOFFICE` / `DEVTOOLS_TESSERACT` / `DEVTOOLS_DRAWIO` に実行ファイルのパスを設定します。

## 対応形式と、実務の文書で対応していること

| 入力 | 対応 |
|---|---|
| `.docx` | 見出し・段落（太字/斜体）・**ハイパーリンク**・箇条書き（レベル）・表（**縦結合・横結合・入れ子の表**）・画像・**テキストボックス**・**脚注/文末脚注**・**コメント**・**変更履歴**（承認後/変更前/差分表示）・**コンテンツコントロール**・**ヘッダー/フッター**・数式（テキスト） |
| `.xlsx` `.xlsm` | **タイトル行・作成者行**を表と区別、**見出し行の自動検出**（`--header-row` で指定も可）、**2 段見出し**（`金額/税抜`）、**セル結合**、**1 シートに複数の表**、**注記行**、**非表示の行・列・シート**（`--skip-hidden`）、**数式の計算結果が無いブックの検知と再計算**（`--recalc`）、パーセント表示、**大きなブックの省メモリ読み込み** |
| `.pptx` | スライドごとの見出し・箇条書き（レベル）・表・画像・発表者ノート・**SmartArt の文字** |
| `.pdf` | 文字サイズ・太字・「第N条」からの見出し推定・箇条書き・表（罫線あり/なし）・**2 段組み**・**繰り返しヘッダー/フッター/ページ番号の除去**・**スキャンページの OCR** |
| `.drawio` | 図形ラベルと接続の一覧＋ソースのコピー＋svg（draw.io がある場合）。圧縮形式にも対応 |
| `.doc` `.xls` `.ppt` `.odt` `.rtf` など | LibreOffice で PDF 化して抽出 |

出力の詳しい約束（見出し・表・脚注の書き方、レポートの rule_id）は [docs/OUTPUT_SPEC.md](docs/OUTPUT_SPEC.md)。

## 使い方

```bash
doc2md 仕様書.docx -d out/                          # out/仕様書.md と out/仕様書.assets/
doc2md docs/ -d out/ -r --continue-on-error         # フォルダ一括（1 ファイル失敗しても続行）
doc2md 見積.xlsx -d out/ --header-row "4月=4"        # 見出し行を指定（全シートなら --header-row 4）
doc2md 台帳.xlsx -d out/ --skip-hidden --recalc      # 非表示を除く・数式を再計算
doc2md 規程.docx -d out/ --track-changes show        # 変更履歴を ~~削除~~ <ins>挿入</ins> で表示
doc2md scan.pdf -d out/ --ocr-lang jpn               # スキャン PDF
doc2md 報告書.pdf -d out/ -f json -o report.json     # 変換レポートを JSON で（CI・AI エージェント用）
doc2md --check-env                                   # 外部ソフトの有無
doc2md docs/ --dry-run                               # 変換せず、各ファイルの読み方と必要なソフトだけ確認
```

### AI に読ませるとき（`--front-matter` / `--llms-txt`）

```bash
doc2md 設計書/ -d out/ -r --front-matter --llms-txt
```

- `--front-matter`: 各 Markdown の先頭に `title` / `id` / `source` / `source_sha256` / `generator` を付けます。
  `id`（例: `DOC-API-001`）は人が付ける識別子で、再変換しても前回の値を引き継ぎます。
  `source_sha256` は元ファイルの内容ハッシュで、「元の設計書が更新されたのに Markdown が古いまま」を検知できます
  （[spec-trace](https://github.com/shoyo-maya-bot/ai-coding/tree/main/tools/spec-trace)（ai-coding に同梱）の設計書ドリフト検知と組み合わせられます）。
- `--llms-txt`: 出力先に `llms.txt`（文書一覧）と `llms-full.txt`（全文を 1 ファイルに連結）を作ります。
  AI に「この資料一式を前提に答えて」と渡すときに使います。

### 図を作る（figspec）

箱と矢印を YAML で書き、**編集できる draw.io ファイル** を作ります。AI に図を描かせるときも、画像ではなくこの形式で出させると、
人が draw.io で手直しでき、Git で差分も追えます。Copilot Chat では `/figspec-draw`。

```bash
figspec check samples/figspec/approval-flow.yaml            # 定義の誤り・線のルール違反
figspec build samples/figspec/approval-flow.yaml --svg      # .drawio（draw.io があれば .svg も）
```

線のルール: 直交線のみ・宙に浮いた線なし・自己ループなし・判断（decision）から出る線にはラベル。
作った .drawio は doc2md で Markdown（要素と接続の一覧）に戻せるので、図の内容をレビューや AI への入力にも使えます。

設計書を Markdown に移す進め方（ロット運用・突合チェックリスト）は [docs/MIGRATION_GUIDE.md](docs/MIGRATION_GUIDE.md)。

### 変換ロスの検知（2 段構え）

1. **構造**: 変換時に認識した見出し・箇条書き・表・行・画像の数と、出力の数を比べる（整形で落ちたものを検知）
2. **本文量**: 元ファイルから **アダプタとは別の経路** で直接数えた文字数（Word/PowerPoint は XML の全テキスト、
   Excel は全セル、PDF は全文字）と、出力の文字数を比べる（未対応の要素で落ちたものを検知 → `loss.raw_text`）

どちらも warning として報告し、`--strict` で終了コード 1 にできます。

### 終了コード

`0` 成功 ／ `1` 変換失敗あり（`--strict` なら警告ありも）／ `2` 引数・入力の誤り、必要な外部ソフトが無い ／ `3` 実行時エラー

## 性能の目安

`python scripts/bench.py` の結果（Linux・Python 3.11。PC の性能で変わります）:

| 入力 | 所要時間 | 最大メモリ |
|---|---|---|
| Excel 100,000 行 × 10 列（100 万セル。既定で省メモリ読み込み） | 約 19 秒 | 約 85 MB |
| Excel 同上を通常読み込み（`--max-cells` を大きくした場合） | 約 26 秒 | 約 860 MB |
| PDF 300 ページ | 約 22 秒 | 約 90 MB |
| Word 5,000 段落 + 表 100 個 | 約 0.4 秒 | 約 95 MB |

Excel は 20 万セル（`--max-cells`）を超えると省メモリ読み込みに切り替わります（結合・複数表・非表示の判定をしない代わりにメモリ一定）。

## 分かっている制限

- PDF の見出し・段落の区切りは推定です。複雑なレイアウト（3 段組み・表と本文の回り込み）は崩れることがあります
- 画像の中の文字（図・スクリーンショット）は OCR しません（スキャンページ全体のみ）
- Excel のグラフ・図形、PowerPoint のグラフ、Word の数式の書式は対象外（文字だけ）
- 「方眼紙」形式の Excel（帳票レイアウト）は表として認識されず、段落の並びになります

## 設計

```mermaid
flowchart LR
  IN[入力ファイル] --> PRE{前処理}
  PRE -- 旧形式 / --via-pdf --> LO[LibreOffice headless<br/>一時ディレクトリで PDF 化] --> PDF[pdf adapter]
  PRE -- --recalc --> RC[LibreOffice で再計算] --> AD
  PRE -- 直接 --> AD[形式別 adapter<br/>docx / xlsx / pptx / pdf / drawio]
  AD --> IR[中間表現（逐次生成）]
  PDF --> IR
  IR --> N[正規化] --> R[Markdown に逐次書き出し<br/>決定的]
  R --> OUT[.md + .assets/]
  AD -. 認識した統計 .-> A[ロス検知]
  IN -. 別経路の文字数 .-> A
  R -. 出力の統計 .-> A
  A --> REP[変換レポート]
```

- **決定的な出力**: 同じ入力・オプションなら同じ Markdown・同じファイル名の抽出図。再実行時は古い `.assets/` を削除
- **逐次処理**: アダプタはブロックを 1 つずつ生成し、正規化・書き出しも逐次（PDF は 1 ページずつキャッシュを解放）
- **機密配慮**: 中間ファイルは `secure_tempdir`（0700・例外時も削除）。ログには件数・ファイル名のみ。ネットワークに接続しない
- **拡張**: 形式を足すときは `adapters/` にモジュールを追加し、`adapters/__init__.py` の `ADAPTERS` に登録（[AGENTS.md](AGENTS.md)）

## 開発

```bash
./setup.sh --dev          # Windows: .\scripts\setup.ps1 -Dev
ruff check . && ruff format --check .
pytest                    # 外部ソフトが無いテストは skip
UPDATE_GOLDEN=1 pytest    # 変換結果を意図して変えたとき期待値を更新（差分を PR でレビュー）
python scripts/bench.py   # 性能測定（要 reportlab）
python scripts/verify_gates.py   # CI と同じゲートを同じ順で（push 前に）
```

テストデータ: `samples/`（プログラム生成）と `samples/realworld/`（**LibreOffice に保存させた** 実務に近い Word/Excel、
2 段組み・スキャン PDF。`samples/realworld/make_realworld.py` で再生成）。内容はすべて架空です。
