# Copilot への指示（doc-to-markdown）

このリポジトリは、Word / Excel / PowerPoint / PDF / draw.io を Markdown に変換する CLI `doc2md` です。
利用者は AI・プログラミングの初心者の場合があります。説明は日本語で、専門用語を避け、手順は番号付きで示してください。

## ツールを使うとき

- 文書の中身を読む必要があるときは、ファイルを直接開かず、ターミナルで次を実行してから Markdown を読む:
  `doc2md "<ファイル>" -d out -f json`（仮想環境が有効でなければ `.venv/Scripts/doc2md`（Windows）/ `.venv/bin/doc2md`）
- 実行結果（標準出力の JSON）の `findings` を必ず確認する:
  - `severity: warning` は利用者に **そのまま伝える**（変更履歴・非表示行・OCR・変換ロスなど、内容が原本と違う可能性）
  - `loss.*` があれば、どこが欠けたかを `--via-pdf` の結果と比べて説明する
- 終了コード 2 は「使い方・入力の誤り」。メッセージを読んで引数を直す。LibreOffice / Tesseract が無いと言われたら
  `doc2md --check-env` の結果を見せ、インストール方法（無料）を案内する
- 変換結果（`out/`）はコミットしない（`.gitignore` 済み）。機密文書の内容をチャットに長く貼らない。要約や該当箇所の引用にとどめる

## コードを変更するとき

- [AGENTS.md](../AGENTS.md) の規約に従う（責務は変換だけ・出力は決定的・本文をログに出さない）
- 変更後: `ruff check . && ruff format --check . && pytest`。変換結果が変わるなら `UPDATE_GOLDEN=1 pytest` で期待値を更新し、差分を説明する
- 出力形式を変えるときは [docs/OUTPUT_SPEC.md](../docs/OUTPUT_SPEC.md) も更新する
