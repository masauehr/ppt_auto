# ppt_auto マニュアル

## 概要

定型PowerPointスライドをPythonまたはVBAマクロで自動生成するツール。
テキスト設定ファイル（JSON）から表紙・目次・本文・まとめスライドを一括生成する。

`python-pptx` で組み立てるため、出力される PPTX は**本物のテキストボックス**を持ち、PowerPoint 上でそのまま編集できる。Marp CLI（[marp_slides](../marp_slides/)）はデザイン自由度が高い一方で PPTX をスライド全体の画像として焼き込むため編集不可という制約があり、その代替として ppt_auto を使う場面がある。比較表は [marp-slides マニュアル「PPTX の編集可否について」](../pc_docs/manuals/automation/marp-slides.md) を参照。

---

## 作業の流れ（はじめての方向け）

基本は「**config.jsonを書く → 生成する → PowerPointで見た目を確認する → 直す**」のサイクルを回すだけ。一発で完璧な見た目にはならないことが多いので、必ず目視確認する前提で作業する。

1. **内容を決めて config.json を書く**（`config.json` を直接編集するか、コピーして新しいファイル名で作る。例: `my_slide_config.json`）
   - 全体設定（`title`/`subtitle`/`author`/`date`）と、`slides` 配列にスライドごとの内容を並べる
   - 詳しいフィールドは下記「設定ファイル（config.json）」参照
2. **生成する**
   ```bash
   cd ~/projects/ppt_auto
   python3 make_pptx.py my_slide_config.json
   ```
   → 同じディレクトリに `my_slide_config.pptx` が生成される（ファイル名は config のファイル名がそのまま使われる）
3. **PowerPointで開いて目視確認する**
   - 文字がボックスからはみ出して他の要素と重なっていないか
   - 画像のサイズ・並びが不自然でないか
   - 誤字・不要な表現が残っていないか
   - **PowerPointで既に開いたまま再生成すると、上書きされたファイルの中身が反映されない（開いているウィンドウが古い内容のまま）ことがある。一度閉じてから開き直すか、再生成前にPowerPoint側を閉じておく**
4. **問題があれば直して再生成**
   - 文言・改行位置の問題 → `config.json` を直す
   - レイアウトそのものの問題（文字が重なる・画像サイズがおかしい等） → `make_pptx.py` 側のスライド生成関数（`make_title_slide` / `make_content_slide` 等）の位置・サイズ指定を直す
   - 直したら手順2に戻って再生成 → 手順3で再確認、を見た目が問題なくなるまで繰り返す

### よくあるつまずきポイント

| 症状 | 原因・対処 |
|------|-----------|
| 箇条書きが改行されず1行に繋がって表示される | （2026-07-01に修正済み）`body`/`quote` 内の `\n` はPowerPoint上では単なる文字であり、そのままでは改行として描画されない。`add_text_box` 側で `\n` ごとに別段落を作るよう対応済みなので、現在のバージョンでは通常発生しない |
| タイトルなど長い文字列が下の要素と重なる | 自動折返しで2行以上になった際にテキストボックスの高さが足りないと起こる。`title` に `\n` を入れて明示的に改行位置を指定するか、`make_pptx.py` 側でボックス高さを広げる |
| `images` で複数画像を並べたとき高さがバラバラに見える | （2026-07-01に修正済み）アスペクト比が画像ごとに異なると、単純な均等幅配分では高さが揃わなかった。現在は全画像の高さを自動で揃えるようになっている |
| 生成したはずなのに古い内容のまま | PowerPointでファイルを開いたまま再生成すると、開いている側は自動更新されない。閉じてから開き直す |

---

## ファイル構成

```
ppt_auto/
├── make_pptx.py             # Python版：config.jsonからPPTX生成
├── make_pptx_from_md.py     # Python版：レイアウト画像＋MarkdownからPPTX生成（AI解析）
├── config.json              # スライド内容の設定ファイル（make_pptx.py用）
├── sample_content.md        # Markdownサンプル（make_pptx_from_md.py用）
└── vba/
    └── make_pptx.bas        # VBAマクロ版
```

---

## Python版

### 必要ライブラリ

```bash
pip install python-pptx
```

### 設定ファイル（config.json）

```json
{
  "title": "プレゼンテーションタイトル",
  "subtitle": "サブタイトル",
  "author": "作成者名",
  "slides": [
    {
      "type": "content",
      "title": "スライドタイトル",
      "body": "本文テキスト",
      "images": ["path/to/image1.png", "path/to/image2.png"],
      "quote": "強調したい一言（左アクセント線付きの帯で表示）"
    }
  ]
}
```

| フィールド | 必須 | 説明 |
|-----------|------|------|
| `type` | ✕（省略時 `content`） | `content`（本文） or `summary`（まとめ・背景強調） |
| `title` | ○ | スライドタイトル。`\n` で改行できる（表紙の `title` が長くて折返しがサブタイトルと重なる場合などに使う） |
| `body` | ✕ | 本文テキスト（`\n` で改行、箇条書きは `・` を先頭に付ける） |
| `image` | ✕ | 画像1枚を挿入（パスは `make_pptx.py` 実行ディレクトリからの相対パス） |
| `images` | ✕ | 画像を複数横並びで挿入（`image` と併用不可）。アスペクト比を保ったまま自動で高さを揃えて配置 |
| `quote` | ✕ | 引用・強調テキストを左アクセント線付きの帯で表示 |

### 実行方法

```bash
cd ~/projects/ppt_auto
python3 make_pptx.py config.json
```

出力ファイル: `<config.jsonのファイル名>.pptx`（スクリプトと同じディレクトリ。例: `ai_projects_intro_config.json` → `ai_projects_intro_config.pptx`）

### 使用例: Marp スライドを編集可能な PPTX に作り直す

Marp CLI（[marp_slides](../marp_slides/)）で作った資料は PPTX がテキスト編集不可（画像焼き込み）なため、PowerPoint 上で手直ししたい場合は本ツールで config.json を作って作り直す。

`ai_projects_intro_config.json` は [marp_slides/slides/2026-07-01-ai-projects-intro.md](../marp_slides/slides/2026-07-01-ai-projects-intro.md) の内容を移植した実例（`image`/`images`/`quote` フィールドの使用例も含む）。

---

## AI解析版（レイアウト画像 + Markdown）

### 概要

レイアウトのお手本画像（PNG/JPG等）とMarkdownファイルを渡すと、Claude APIが画像からデザインを解析して、そのデザインでPPTXを自動生成します。

### 必要なもの

- `ANTHROPIC_API_KEY` 環境変数（Claude APIキー）
- `anthropic` ライブラリ（`pip install anthropic`）

### 実行方法

```bash
export ANTHROPIC_API_KEY="your-api-key"
python3 make_pptx_from_md.py <レイアウト画像> <Markdownファイル> [出力.pptx]
```

```bash
# 例: layout.png + sample_content.md → sample_content.pptx
python3 make_pptx_from_md.py layout.png sample_content.md

# 出力先を指定する場合
python3 make_pptx_from_md.py layout.png sample_content.md output/my_slide.pptx
```

### Markdownフォーマット仕様

```markdown
---
title: プレゼンテーションタイトル
subtitle: サブタイトル
author: 作成者名
date: 2026年4月29日
---

## スライドタイトル1

本文テキスト
- 箇条書き1（→ • に変換）
- 箇条書き2

## スライドタイトル2

1. 番号付きリスト
2. 番号付きリスト

## まとめ <!-- summary -->

まとめの内容（背景が強調色になる）
```

| 記法 | 意味 |
|------|------|
| `---`...`---` | フロントマター（タイトル・作成者・日付） |
| `## タイトル` | スライド区切り＆タイトル |
| `- テキスト` | 箇条書き（• に変換） |
| `<!-- summary -->` | まとめスライドとしてマーク |

### 動作の流れ

1. **レイアウト解析**: Claude APIがお手本画像を解析し、背景色・ヘッダー色・フォントサイズ等を抽出
2. **キャッシュ**: 解析結果は `<画像名>_layout_cache.json` に保存。2回目以降はAPI呼び出しなしで実行
3. **Markdown解析**: フロントマター・H2見出し・本文・箇条書きを構造化
4. **PPTX生成**: 解析したデザインをそのままスライドに適用

---

## VBA版

### セットアップ

1. PowerPointを開く
2. Alt+F11でVBAエディタを起動
3. ファイル → ファイルのインポート で `vba/make_pptx.bas` を選択
4. `MakePPTX` マクロを実行

### 動作

- PowerPointのVBAから直接スライドを生成する
- テキストはコード内の設定部分に直書き（または別途テキストファイルから読み込み）

---

## 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-04-29 | プロジェクト新規作成 |
| 2026-04-29 | `make_pptx_from_md.py` 追加（レイアウト画像＋MarkdownからAI解析でPPTX生成） |
| 2026-07-01 | 概要に Marp CLI との違い（テキスト編集可否）を追記。比較表は marp-slides マニュアル側に集約 |
| 2026-07-01 | `make_pptx.py` に `image`/`images`/`quote` フィールドを追加（画像挿入・引用帯表示に対応）。出力ファイル名を `output.pptx` 固定から `<config名>.pptx` に変更。使用例として `ai_projects_intro_config.json`（Marpスライドの移植）を追加 |
| 2026-07-01 | バグ修正: `add_text_box` が `\n` を1つのrunにそのまま入れていたため箇条書きが改行されず繋がって表示される問題を修正（行ごとに別段落を生成するよう変更）。`add_images_row` をアスペクト比の異なる画像でも高さが揃うロジックに変更。表紙タイトルが長い場合に下のサブタイトルと重なる問題を修正（タイトルボックスを拡大・`\n`で明示的な改行に対応）。あわせて「作業の流れ（はじめての方向け）」セクションを追加 |
