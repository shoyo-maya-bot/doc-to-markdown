---
description: 文書（Word/Excel/PowerPoint/PDF/draw.io）を Markdown に変換し、注意点と中身の概要を報告する
agent: agent
argument-hint: 変換するファイルかフォルダのパス
---

次の文書を Markdown に変換して、結果を日本語で報告してください。

対象: ${input:path:変換するファイルかフォルダのパス}

## 手順

1. ターミナルで実行する（フォルダなら `-r` を付ける）:
   `doc2md "<対象>" -d out -f json`
   - `doc2md` が見つからなければ `.venv/Scripts/doc2md`（Windows）/ `.venv/bin/doc2md`（macOS・Linux）で実行する
   - それでも無ければ、`setup.cmd`（Windows）/ `./setup.sh` を実行するよう案内して止まる
2. 標準出力の JSON を読み、次の形で報告する:
   - **変換したファイル**: `convert.ok` の出力先と件数
   - **確認が必要なこと**: `severity` が `warning` / `error` のものを全部、1 行ずつ分かりやすい言葉で
     （例:「未確定の変更履歴が 4 件あります。承認後の内容で出力しています」）
   - 問題が無ければ「注意点はありません」と書く
3. 出力された Markdown（`out/<名前>.md`）を開き、見出しの構成と主な内容を 5 行以内で要約する
4. 次にできることを 1 つだけ提案する（例: 表の整合性チェックなら fieldcheck、変換ロスがあれば `--via-pdf` で比較）

## 注意

- 元の文書ファイルは直接開かない（バイナリのため）。必ず変換結果を読む
- 機密の可能性がある本文をチャットに長く引用しない
