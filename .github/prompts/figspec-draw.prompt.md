---
description: 説明文から figspec（図の定義 YAML）を書き、編集できる draw.io の図を作る
agent: agent
argument-hint: 描きたい図の説明（例「申請→課長承認→経理。10 万円以上は部長承認を挟む」）
---

次の図を figspec で作ってください。

${input:desc:描きたい図の説明}

## 手順

1. `src/doc_to_markdown/figspec.py` の先頭の説明（定義の書き方と線のルール）を読む
2. `figures/<名前>.yaml` に定義を書く（ノード ID は英小文字、ラベルは日本語、判断は `shape: decision`）
3. `figspec check figures/<名前>.yaml` を実行し、error と warning が 0 になるまで直す
4. `figspec build figures/<名前>.yaml --svg` で draw.io（と、draw.io があれば svg）を作る
5. 作った図の要素と接続を箇条書きで報告し、説明と食い違いがないか確認をお願いする

## 守ること

- 説明に無い要素・分岐を足さない。分からない分岐は質問する
- 線のルール（直交・宙に浮いた線なし・判断の線にラベル）を守る
