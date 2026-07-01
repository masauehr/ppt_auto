# ppt_auto マニュアル

## 概要

定型PowerPointスライドをPythonまたはVBAマクロで自動生成するツール。
テキスト設定ファイル（JSON）から表紙・目次・本文・まとめスライドを一括生成する。

`python-pptx` で組み立てるため、出力される PPTX は**本物のテキストボックス**を持ち、PowerPoint 上でそのまま編集できる。Marp CLI（[marp_slides](../marp_slides/)）はデザイン自由度が高い一方で PPTX をスライド全体の画像として焼き込むため編集不可という制約があり、その代替として ppt_auto を使う場面がある。比較表は [marp-slides マニュアル「PPTX の編集可否について」](../pc_docs/manuals/automation/marp-slides.md) を参照。

**入力形式は2種類**（同じ `make_pptx.py` で、ファイルの拡張子を見て自動で切り替わる）:

- **Markdown（`.md`）** ─ Marpと同じ感覚で書ける簡易記法。まずはこちらを推奨
- **JSON（`.json`）** ─ 画像の細かい配置など、より細かく制御したい場合向け

---

## 作業の流れ（はじめての方向け）

基本は「**Markdownを書く → 生成する → PowerPointで見た目を確認する → 直す**」のサイクルを回すだけ。一発で完璧な見た目にはならないことが多いので、必ず目視確認する前提で作業する。

1. **内容を決めて Markdown を書く**（[sample_slides.md](sample_slides.md) をコピーして書き換えるのが早い。例: `my_slides.md`）
   - 冒頭の `---`〜`---` に `title`/`subtitle`/`author`/`date` を書く
   - `##` でスライドを区切り、本文・箇条書き（`-`）・引用（`>`）・画像（`![説明](パス)`）を書く
   - 詳しい記法は下記「Markdownから生成する」参照
2. **生成する**
   ```bash
   cd ~/projects/ppt_auto
   python3 make_pptx.py my_slides.md
   ```
   → 同じディレクトリに `my_slides.pptx` が生成される（ファイル名は入力ファイルの名前がそのまま使われる）
3. **PowerPointで開いて目視確認する**
   - 文字がボックスからはみ出して他の要素と重なっていないか
   - 画像のサイズ・並びが不自然でないか
   - 誤字・不要な表現が残っていないか
   - **PowerPointで既に開いたまま再生成すると、上書きされたファイルの中身が反映されない（開いているウィンドウが古い内容のまま）ことがある。一度閉じてから開き直すか、再生成前にPowerPoint側を閉じておく**
4. **問題があれば直して再生成**
   - 文言・改行位置の問題 → Markdownファイルを直す
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

## Markdownから生成する（推奨）

`config.json` を書かなくても、Markdownファイル1枚から直接PPTXを生成できる。Marpの記法に近いシンプルなルールのみ。

```bash
cd ~/projects/ppt_auto
python3 make_pptx.py sample_slides.md
```

### 記法

```markdown
---
title: プレゼンテーションタイトル
subtitle: サブタイトル
author: 作成者名
date: 2026年7月
---

## スライドタイトル

本文テキスト
- 箇条書き1（自動で「・」に変換される）
- 箇条書き2

> 引用・強調したい一言（左アクセント線付きの帯で表示される）

![説明](画像のパス)
![説明](画像のパス2)

## まとめ <!-- summary -->

まとめの内容（背景が強調色になる）
```

| 記法 | 意味 |
|------|------|
| `---`...`---` | フロントマター（タイトル・サブタイトル・作成者・日付） |
| `## タイトル` | スライド区切り＆タイトル |
| `- テキスト` | 箇条書き（`・` に変換） |
| `> テキスト` | 引用（quoteフィールドとして左アクセント線付きの帯で表示） |
| `![説明](パス)` | 画像挿入。複数行書くと横並びで挿入され、高さが自動で揃う。パスはMarkdownファイル自身からの相対パス |
| `<!-- summary -->` | まとめスライドとしてマーク（`##` の行末に書く） |

サンプルは [sample_slides.md](sample_slides.md) を参照（生成すると `sample_slides.pptx` ができる）。実例として `ai_projects_intro_config.json`（JSON版）もあるが、同等のものをMarkdownで書く場合は上記記法でよい。

より細かい配置調整（画像の混在パターンなど）をしたい場合は、次項のJSON版（`config.json`）も引き続き使える。

---

## ファイル構成

```
ppt_auto/
├── make_pptx.py             # Python版：Markdown(.md) または config.json(.json) からPPTX生成
├── make_pptx_from_md.py     # Python版：レイアウト画像＋MarkdownからPPTX生成（AI解析。別ツール）
├── sample_slides.md         # Markdownサンプル（make_pptx.py用。まずこれを見る）
├── config.json              # JSON設定サンプル（make_pptx.py用）
├── sample_content.md        # Markdownサンプル（make_pptx_from_md.pyのAI解析版用。別記法）
└── vba/
    └── make_pptx.bas        # VBAマクロ版
```

---

## Python版

### 必要ライブラリ

```bash
pip install python-pptx
```

### JSON版で細かく制御する（config.json）

より複雑なレイアウトを作りたい場合は、Markdownの代わりにJSON形式の `config.json` を直接編集することもできる（`make_pptx.py` はファイルの拡張子で自動判別する）。

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

別ツール `make_pptx_from_md.py` の機能。**上記「Markdownから生成する」（`make_pptx.py`）とは別物**で、Markdownの記法も少し異なる（こちらは `- 箇条書き` が `・` ではなく `•` に変換される、`quote`/`images` フィールドに相当する記法はない）。決まったデザイン（config.json/Markdown）で十分な場合は上記の `make_pptx.py` を使う。手本となるスライド画像がありAPIキーも用意できる場合のみ、こちらを検討する。

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
| 2026-07-01 | `make_pptx.py` がMarkdownファイル（`.md`）を直接読み込めるように対応（`config.json` が不要に）。`##`見出し・`-`箇条書き・`>`引用・`![]()`画像のシンプルな記法に対応し、Marpに近い感覚で使える。サンプル `sample_slides.md` を追加。マニュアルを全面的に「Markdownから生成する」方法を軸にした構成に変更 |
