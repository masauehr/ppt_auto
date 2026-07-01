# ppt_auto — PowerPoint自動作成ツール

詳しくは [ppt_auto_manual.md](ppt_auto_manual.md) を参照。

## 概要

Pythonで定型スライドを自動生成し、VBAマクロ版でもExcelから直接実行できる。Markdownファイル1枚から、Marpに近い感覚でPPTXを生成できる。

## 構成

```
ppt_auto/
├── README.md                # このファイル
├── ppt_auto_manual.md       # 詳細マニュアル
├── make_pptx.py             # Python版：Markdown(.md) または config.json(.json) からPPTX生成
├── make_pptx_from_md.py     # Python版：レイアウト画像＋MarkdownからPPTX生成（AI解析。別ツール）
├── sample_slides.md         # Markdownサンプル（make_pptx.py用）
├── config.json              # JSON設定サンプル（make_pptx.py用）
├── sample_content.md        # Markdownサンプル（make_pptx_from_md.pyのAI解析版用。別記法）
└── vba/
    └── make_pptx.bas        # VBAマクロ版
```

## 使い方①：Markdownから生成（推奨・シンプル）

```bash
cd ~/projects/ppt_auto
python3 make_pptx.py sample_slides.md
```

```markdown
---
title: プレゼンタイトル
subtitle: サブタイトル
author: 作成者名
date: 2026年7月
---

## スライドタイトル

本文テキスト
- 箇条書き1
- 箇条書き2

> 引用・強調したい一言

![説明](画像のパス)

## まとめ <!-- summary -->

まとめ本文
```

## 使い方②：config.json から生成（細かく制御したい場合）

```bash
python3 make_pptx.py config.json
```

## 使い方③：レイアウト画像 + Markdown から生成（AI解析版）

```bash
# ANTHROPIC_API_KEY が必要
python3 make_pptx_from_md.py <レイアウト画像> <Markdownファイル> [出力.pptx]

# 例
python3 make_pptx_from_md.py layout.png sample_content.md
```

## 使い方（VBA版）

1. PowerPointを開き、Alt+F11でVBAエディタを起動
2. `vba/make_pptx.bas` をインポート
3. `MakePPTX` マクロを実行
