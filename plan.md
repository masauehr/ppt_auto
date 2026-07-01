# plan.md — ppt_auto 今後の予定

## 完了

- [x] `ppt_auto_manual.md`（および `pc_docs/manuals/automation/ppt-auto.md`）のVBA版セクションに、Mac版のVBAエディタ起動コマンドを追記する（2026-07-01完了）
  - Windows版「Alt+F11」を残しつつ、Mac版「Option+F11」（F11がMission Control割り当て済みの場合は`Fn+Option+F11`）を併記
  - あわせて `vba/make_pptx.bas` のコンパイルエラー・文字化け対応（Shift-JISへのエンコーディング修正）も実施

## 完了（続き）

- [x] VBA版に `MakePPTXFromMarkdown` マクロを追加（2026-07-01完了）
  - `make_pptx.py` のMarkdown版とほぼ同じ記法（フロントマター・`##`見出し・箇条書き・`<!-- summary -->`）に対応
  - `>`引用・`![]()`画像は非対応（VBA版は文字列処理のみのため簡易記法に限定）
- [x] `MakePPTXFromMarkdown` の動作確認（2026-07-01、実際にPowerPoint(Mac)で実行し最終的に成功）
  - 1回目: スライド生成自体は成功したが、本文が文字化け
    - 原因: Mac版VBAの `Open...For Input`/`Line Input` がUTF-8マルチバイト文字を正しく読めない
    - 対応: `ReadTextFile` をバイナリ読み込み＋自前UTF-8デコード（`Utf8BytesToString`）に変更
  - 2回目: 文字数が本来の2倍になり、内容も破損（"## "見出しも検出できず「スライドが見つかりません」エラー）
    - 原因: `Utf8BytesToString` 内の `If codepoint > &HFFFF Then`（サロゲートペア判定）で使っていた16進数リテラル `&HFFFF`・`&HD800`・`&HDC00` が、VBAでは符号付き16bit Integerとして解釈され負数（`&HFFFF`→-1等）になってしまい、あらゆる文字が常にサロゲートペア分岐に入り1文字が2文字に分裂していた
    - 対応: 該当箇所をすべて10進数リテラル（65535・55296・56320）に置き換え
  - 3回目: 最小テスト（`TestMinimalUtf8Decode`、ファイルI/O抜き）で修正を確認 → 正常
  - 4回目: `MakePPTXFromMarkdown` を `sample_slides.md` で実行 → **成功**（文字化けなし、正しくスライド生成）
  - デバッグ用に追加したコード（`gDebugTrace`グローバル変数、`TestMinimalUtf8Decode`、`DebugDumpRawBytes`、各種トレース出力）はすべて削除し、クリーンな状態に戻し済み

## TODO

（現在なし。次のタスクがあれば追記する）
