# AGENTS.md — AI エージェント向けガイド（doc-to-markdown）

## このツールを「使う」とき

- ドキュメント（.docx/.xlsx/.pptx/.pdf/.drawio ほか）の中身を読む必要があれば、バイナリを直接読まずに
  `doc2md <file> -d <outdir> -f json` で Markdown にしてから `<outdir>/<stem>.md` を読む。
- 標準出力の JSON の `findings` に `loss.*`（warning）があれば、変換で欠けた可能性がある。
  利用者に伝えるか、`--via-pdf` で再変換して比較する。
- 終了コード: 0 成功 / 1 失敗またはロス（--strict）/ 2 使い方の誤り / 3 実行時エラー。
- 検証（「合計が合っているか」等）はこのツールの責務外。表を `fieldcheck validate <md> -r rules.yaml --table <見出し>` に渡す。

## このリポジトリを「変更する」とき

- 責務は「変換」だけ。検証・要約・翻訳などの機能を足さない（別ツールにする）。
- 出力は決定的に保つ（時刻・乱数・辞書順に依存しない順序を入れない）。
- 形式を追加する: `src/doc_to_markdown/adapters/<fmt>_adapter.py` に `convert(path, assets, opts) -> Document` を実装し、
  `adapters/__init__.py` の `ADAPTERS` に登録。次を必ず守る（ロス検知に使う）:
  - 認識した要素の統計（`doc.stats`）を数える
  - 可能なら、アダプタとは別の経路で元ファイルの文字数を数えて `doc.raw_text_chars` に入れる
  - 変換側が付け足す見出し等は `synthetic=True` にする
  - 大きなファイルに備えて `doc.blocks` はジェネレータで返してよい（stats は最後まで読んだ時点で確定）
- 利用者が確認すべきこと（変更履歴・非表示・OCR など）は `doc.warnings`、補足は `doc.notes` に入れる
- 外部ソフトは `devtools_common.executables` で探す（`shutil.which` を直接使わない。Windows の標準インストール先を見ないため）
- 出力の約束を変えたら [docs/OUTPUT_SPEC.md](docs/OUTPUT_SPEC.md) を更新する
- 変更後に必ず実行: `ruff check . && ruff format --check . && pytest`
- 変換結果が変わるのが意図どおりなら `UPDATE_GOLDEN=1 pytest` → `tests/golden/` の差分を確認して PR に含める。
- 機密: 本文をログに出さない（`devtools_common.log.describe_text` を使う）。一時ファイルは `secure_tempdir`。
- 共通処理（CLI 規約・レポート・ログ・一時ファイル）は `devtools-common` にある。ここへコピーしない。
