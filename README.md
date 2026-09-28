# doc-to-markdown（doc2md）

各種ドキュメント（Office / PDF / 図）を **Markdown + 抽出図** に変換することだけに責務を絞ったツールです。
内容の検証（項目間の整合性チェック）は行いません → [field-validation](https://github.com/<your-org>/field-validation) に別プロセスで渡します。

用途の例: 仕様書・手順書を Git 管理下のテキストにする／生成 AI に読ませる前処理／差分レビュー／検索インデックス化

## 対応形式

| 入力 | 方式 | 出力されるもの |
|---|---|---|
| `.docx` | python-docx | 見出し階層・段落（太字/斜体）・箇条書き（番号/記号・レベル）・表・画像 |
| `.xlsx` `.xlsm` | openpyxl（ストリーム） | シートごとに見出し＋表（数式は計算済みの値） |
| `.pptx` | python-pptx | スライドごとに見出し・箇条書き（レベル）・表・画像・発表者ノート |
| `.pdf` | pdfplumber | 文字サイズから見出し推定・箇条書き・表（罫線あり / 罫線なしの列揃え）・ページ番号コメント |
| `.drawio` | XML 解析（圧縮形式対応） | 図形ラベルと接続の一覧＋ソースのコピー＋svg（draw.io CLI がある場合） |
| `.doc` `.xls` `.ppt` `.odt` `.rtf` など | LibreOffice headless で PDF 化 → PDF として抽出 | 同上（PDF） |

`--via-pdf` を付けると `.docx` / `.pptx` なども PDF 経由で抽出します（レイアウト優先）。
Excel は PDF 経由だと列幅不足で値が欠けやすいため、直接変換（既定）を推奨します。

## インストール

```bash
# 利用者（CLI として入れる）
pipx install "git+https://github.com/<your-org>/doc-to-markdown.git@v0.1.0"

# 開発者
git clone https://github.com/<your-org>/doc-to-markdown.git && cd doc-to-markdown
python -m venv .venv && source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
```

任意の外部ツール: LibreOffice（`--via-pdf` と旧形式）、draw.io デスクトップ（svg 書き出し）。

## 使い方

```bash
doc2md 仕様書.docx -d out/                         # out/仕様書.md と out/仕様書.assets/
doc2md docs/ -d out/ --recursive                  # フォルダ一括（バッチ）
doc2md docs/ -d out/ -r --continue-on-error       # 1 ファイル失敗しても続行
doc2md 見積.pdf -d out/ --format json -o report.json   # 変換レポートを JSON で（CI・AI エージェント用）
doc2md 手順書.docx -d out/ --via-pdf --strict      # PDF 経由。ロスの疑いがあれば終了コード 1
```

出力:

```text
out/
├── 仕様書.md
└── 仕様書.assets/          # 抽出図（<stem>-01.png …）。draw.io はソース(.drawio)と svg をペアで保管
    └── 仕様書-01.png
```

### 変換レポート（標準出力）

変換結果の統計と **変換ロスの疑い** を出します（Markdown / HTML / CSV / JSON）。

| rule_id | 重大度 | 内容 |
|---|---|---|
| `convert.ok` | info | 出力先と元文書の統計（見出し・箇条書き・表・行・画像・文字数） |
| `loss.headings` / `loss.list_items` / `loss.tables` / `loss.table_rows` / `loss.images` | warning | 元文書と出力で数が一致しない |
| `loss.text` | warning | 本文の非空白文字数が元の 95% 未満 |
| `convert.note` | info | 形式ごとの注意事項（対象外の要素など） |
| `convert.failed` | error | 変換失敗（`--continue-on-error` 時） |

### 終了コード

`0` 成功 ／ `1` 変換失敗あり、または `--strict` 時にロスの疑いあり ／ `2` 引数・入力の誤り ／ `3` 実行時エラー

## 設計

```mermaid
flowchart LR
  IN[入力ファイル] --> PRE{前処理}
  PRE -- 旧形式 / --via-pdf --> LO[LibreOffice headless<br/>一時ディレクトリで PDF 化] --> PDF[pdf adapter]
  PRE -- 直接 --> AD[形式別 adapter<br/>docx / xlsx / pptx / pdf / drawio]
  AD --> IR[中間表現<br/>Heading / Paragraph / ListItem / Table / Image]
  PDF --> IR
  IR --> N[正規化<br/>見出し階層・表・箇条書き]
  N --> R[Markdown 描画<br/>決定的]
  R --> OUT[.md + .assets/]
  IR -. 元文書の統計 .-> A[ロス検知]
  R -. 出力の統計 .-> A
  A --> REP[変換レポート]
```

- **決定的な出力**: 同じ入力からは常に同じ Markdown・同じファイル名の抽出図（タイムスタンプ等を含めない）。
  ゴールデンテストと差分レビューが成り立つ。再実行時は古い `.assets/` を削除して作り直す。
- **正規化ルール**: 見出しは 1 段ずつ深く（H1→H3 は H1→H2）、空要素の除去、空白の統一、表は矩形化し空列を削除、
  箇条書きのレベルは直前 +1 まで。
- **機密配慮**: 中間 PDF は `secure_tempdir`（0700・例外時も削除）で扱う。ログには件数・ファイル名のみで本文を出さない。
- **拡張**: 形式を足すときは `adapters/` にモジュールを追加し、`adapters/__init__.py` の `ADAPTERS` に拡張子を登録する。

## 開発

```bash
ruff check . && ruff format --check .
pytest                              # または python -m unittest discover -s tests
UPDATE_GOLDEN=1 pytest              # 変換結果を意図して変えたとき、期待値を更新（差分を PR でレビュー）
pip install reportlab && python samples/make_samples.py   # サンプル文書の再生成（通常は不要）
```

`samples/` の文書と `tests/golden/` の期待値はすべて架空のデータです。
