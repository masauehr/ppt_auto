# ppt_auto マニュアル

## 概要

定型PowerPointスライドをPythonまたはVBAマクロで自動生成するツール。
テキスト設定ファイル（JSON）から表紙・目次・本文・まとめスライドを一括生成する。

`python-pptx` で組み立てるため、出力される PPTX は**本物のテキストボックス**を持ち、PowerPoint 上でそのまま編集できる。Marp CLI（[marp_slides](https://github.com/masauehr/marp_slides)）はデザイン自由度が高い一方で PPTX をスライド全体の画像として焼き込むため編集不可という制約があり、その代替として ppt_auto を使う場面がある。比較表は [marp-slides マニュアル「PPTX の編集可否について」](https://github.com/masauehr/pc_docs/blob/main/manuals/automation/marp-slides.md) を参照。

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

Marp CLI（[marp_slides](https://github.com/masauehr/marp_slides)）で作った資料は PPTX がテキスト編集不可（画像焼き込み）なため、PowerPoint 上で手直ししたい場合は本ツールで config.json を作って作り直す。

`ai_projects_intro_config.json` は [marp_slides/slides/2026-07-01-ai-projects-intro.md](https://github.com/masauehr/marp_slides/blob/main/slides/2026-07-01-ai-projects-intro.md) の内容を移植した実例（`image`/`images`/`quote` フィールドの使用例も含む）。

---

## AI解析版（レイアウト画像 + Markdown）

別ツール `make_pptx_from_md.py` の機能。**上記「Markdownから生成する」（`make_pptx.py`）とは別物**で、Markdownの記法も少し異なる（こちらは `- 箇条書き` が `・` ではなく `•` に変換される、`quote`/`images` フィールドに相当する記法はない）。決まったデザイン（config.json/Markdown）で十分な場合は上記の `make_pptx.py` を使う。手本となるスライド画像がありAPIキーも用意できる場合のみ、こちらを検討する。

### 概要

レイアウトのお手本画像（PNG/JPG等）とMarkdownファイルを渡すと、Claude APIが画像からデザインを解析して、そのデザインでPPTXを自動生成します。

**ここで言う「Claude API」は、Anthropic社が提供するAPI（`anthropic` Pythonライブラリ経由でのAPI呼び出し）のことで、Claude Code（このCLIツール自体）とは別物**。`ppt_auto` の他のスクリプト（`make_pptx.py` のMarkdown版・JSON版、VBA版）はAIを一切使わず `python-pptx` のみでスライドを組み立てるが、この `make_pptx_from_md.py` だけは以下の処理でClaude APIを呼び出す。

- 呼び出しタイミング: 渡したレイアウトのお手本画像（PNG/JPG）を解析するときのみ（`analyze_layout()` 関数）。背景色・ヘッダー色・フォントサイズなど、デザイン設定をJSON形式で抽出させる
- 利用モデル: Claude Sonnet（Vision機能を使用）
- Markdownの内容そのものの解釈・スライド生成にはAIを使わない（プレーンなテキスト処理）
- 解析結果は `<画像名>_layout_cache.json` にキャッシュされるため、同じ画像であれば2回目以降はAPI呼び出しが発生しない
- API利用には [Anthropic Console](https://console.anthropic.com/) で発行する `ANTHROPIC_API_KEY` が必要。Claude Codeのサブスクリプションとは別に、APIの利用量に応じた従量課金が発生する

### 必要なもの

- `ANTHROPIC_API_KEY` 環境変数（Anthropic Consoleで発行するClaude APIキー）
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

VBAモジュール（`vba/make_pptx.bas`）には2つのマクロが収録されている。

| マクロ名 | 内容 |
|---|---|
| `MakePPTX` | コード内に直書きした雛形（`titles`/`bodies`/`types` 配列）からスライドを生成 |
| `MakePPTXFromMarkdown` | Markdownファイルを読み込んでスライドを生成（`make_pptx.py` と同じ簡易記法に対応） |

### セットアップ

1. PowerPointを開く
2. VBAエディタを起動（Windows: `Alt+F11` / Mac: `Option+F11`。macOSでF11がMission Control等に割り当てられている場合は `Fn+Option+F11`）
3. ファイル → ファイルのインポート で `vba/make_pptx.bas` を選択
4. `MakePPTX` または `MakePPTXFromMarkdown` マクロを実行

### `MakePPTXFromMarkdown`（Markdownから生成）の使い方

1. マクロを実行すると、Markdownファイルのフルパスを入力するダイアログが表示される（Mac版VBAは `Application.FileDialog` が使えないため、ファイル選択ダイアログではなくパス直接入力方式）
2. 例: `/Users/username/projects/ppt_auto/sample_slides.md`
3. 対応する記法は `make_pptx.py` のMarkdown版とほぼ同じ（フロントマターの `title`/`subtitle`/`author`/`date`、`##` 見出し、`-`/`*` 箇条書き、`<!-- summary -->`）
4. **`make_pptx.py` との違い**: `>` 引用（quote）と `![]()` 画像挿入には対応していない（VBA版は文字列処理のみで実装しているため、シンプルな記法に絞っている）。画像・引用を使いたい場合は `make_pptx.py`（Python版）を使う

**Markdownファイル自体の文字コードはUTF-8で保存すること**（`sample_slides.md` 等、このプロジェクトの.mdファイルは通常UTF-8）。Mac版VBAの `Open...For Input`/`Line Input` はUTF-8のマルチバイト文字を正しく読み込めず文字化けするため、`ReadTextFile` 関数はテキストとしてではなくバイナリで読み込み、自前でUTF-8をデコードする実装になっている（`vba/make_pptx.bas` の `.bas` ファイル自体をShift-JISで保存しているのとは別問題で混同しないよう注意）。

**開発メモ（VBAの罠）**: `Utf8BytesToString` 内のUTF-8デコード処理で、サロゲートペア判定に `&HFFFF`・`&HD800`・`&HDC00` のような16進数リテラルを使うと、VBAはこれらを**符号付き16bit Integer**として解釈してしまい負数になる（例: `&HFFFF` → `-1`）。その結果 `If codepoint > &HFFFF Then` が常に真になり、あらゆる文字（ASCII含む）が誤ってサロゲートペア扱いされ、文字数が2倍になり内容が破損するというバグが実際に発生した。回避策として該当箇所はすべて10進数リテラル（`65535`・`55296`・`56320`）を使っている。大きな16進数リテラルをVBAで扱う際は一般的に注意が必要（Long型として扱いたい場合は末尾に `&` を付けるか、10進数で書く）。

### 動作

- PowerPointのVBAから直接スライドを生成する
- `MakePPTX`: テキストはコード内の設定部分に直書き
- `MakePPTXFromMarkdown`: テキストは外部のMarkdownファイルから読み込む

### 文字化け・コンパイルエラーが出る場合

`vba/make_pptx.bas` は **Shift-JIS** エンコーディングで保存されている（VBAエディタの「ファイルのインポート」がUTF-8を正しく解釈できず、日本語部分が文字化けした上でコンパイルエラー（構文エラー）になるため）。

- GitHub上でこのファイルを直接見ると文字化けして見えるが、これは想定どおり（Shift-JISのバイト列をUTF-8として表示しているため）。ダウンロードしてVBAエディタにインポートすれば正しく表示される
- 手元で編集する場合は、保存時にエンコーディングをShift-JISのまま維持すること（UTF-8で保存し直すと再び文字化け・コンパイルエラーの原因になる）

文字化けが解消した後に「コンパイルエラーです。: Sub または Function が定義されていません。」／「メソッドまたはデータ メンバーが見つかりません。」（いずれも `CentimetersToPoints` の箇所を指す）が出る場合は、**Mac版PowerPointのVBAには `Application.CentimetersToPoints` 自体が実装されていない**ことが原因（Windows版には存在するがMac版のVBAランタイムには存在しない）。現在のバージョンでは、この関数に依存せず自前の変換関数 `CmToPt(cm As Single)`（`cm * 28.346456692913385` で計算。1cm = 72/2.54 pt）に置き換え済みのため、最新の `vba/make_pptx.bas` を使えば発生しない。

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
| 2026-07-01 | README/マニュアルの誤記修正: VBAマクロ版は「Excelから実行」ではなく「PowerPointのVBAエディタから実行」が正しいため訂正。あわせて「AI解析版（`make_pptx_from_md.py`）のみがClaude API（Anthropic APIのことでClaude Codeとは別物）を、レイアウト画像解析にのみ使用する」という利用範囲を詳述 |
| 2026-07-01 | README.mdの概要文から誤解を招く「（Excelではなく...）」の補足を削除（VBA版の説明は既に本文で正しく記載されているため重複・冗長だった）。利用モデルの表記を `claude-sonnet-4-6` のような特定バージョン番号付きから「Claude Sonnet」という総称表記に変更（モデルは更新されていくため、ドキュメントにバージョン番号を固定しない方針） |
| 2026-07-01 | `make_pptx_from_md.py` のレイアウト解析で呼び出すモデルを `claude-sonnet-4-6` から最新の `claude-sonnet-5` に更新 |
| 2026-07-01 | `marp_slides` プロジェクトへのリンクが `../marp_slides/` という相対パスになっており、`ppt_auto` が独立したGitHubリポジトリになったことでリンク切れになっていた問題を修正。`https://github.com/masauehr/marp_slides` への絶対URLに変更 |
| 2026-07-01 | `vba/make_pptx.bas` がMacのVBAエディタでインポート時に文字化け・コンパイルエラーになる問題を修正（UTF-8で保存されていたためVBAエディタがShift-JISとして誤解釈していた。Shift-JIS＋CRLFに変換）。`.gitattributes` で `vba/*.bas` を改行コード変換対象外に設定。VBA版セットアップ手順にMac版起動コマンド（`Option+F11`）を追記 |
| 2026-07-01 | `CentimetersToPoints` 関連のコンパイルエラーを修正。`Application.CentimetersToPoints(...)` への明示修飾でも「メソッドまたはデータ メンバーが見つかりません」というエラーが再発したため、Mac版PowerPointのVBAには本メソッド自体が実装されていないと判断。自前の変換関数 `CmToPt(cm As Single)`（`cm * 28.346456692913385`）を追加し、全37箇所を置き換え |
| 2026-07-01 | VBA版に `MakePPTXFromMarkdown` マクロを追加。`make_pptx.py` のMarkdown版とほぼ同じ記法（フロントマター・`##`見出し・箇条書き・`<!-- summary -->`）に対応し、外部Markdownファイルからスライドを生成できるようにした（`>`引用・`![]()`画像は非対応）。ファイル選択はMac版VBAで`Application.FileDialog`が使えないためInputBoxでパス入力方式とした |
| 2026-07-01 | `MakePPTXFromMarkdown` 実行時にMarkdownの内容が文字化けする問題を修正。Mac版VBAの `Open...For Input`/`Line Input` はUTF-8のマルチバイト文字を正しく読み込めないため、`ReadTextFile` をバイナリ読み込み＋自前UTF-8デコード（`Utf8BytesToString` 関数を新規追加。BOM除去・サロゲートペア対応込み）に変更 |
| 2026-07-01 | `Utf8BytesToString` のバグを修正: サロゲートペア判定の16進数リテラル（`&HFFFF`・`&HD800`・`&HDC00`）がVBAで符号付きIntegerとして解釈され負数になっていたため、あらゆる文字が誤って2文字に分裂して文字化けしていた。10進数リテラルに置き換えて解消。`sample_slides.md` で正常に日本語表示されることを実機（Mac版PowerPoint）で確認済み |
