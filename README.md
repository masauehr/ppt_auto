# ppt_auto — PowerPoint自動作成ツール

詳しくは [ppt_auto_manual.md](ppt_auto_manual.md) を参照。

## 概要

Pythonで定型スライドを自動生成し、VBAマクロ版でもExcelから直接実行できる。

## 構成

```
ppt_auto/
├── README.md                # このファイル
├── ppt_auto_manual.md       # 詳細マニュアル
├── make_pptx.py             # Python版：config.jsonからPPTX生成
├── make_pptx_from_md.py     # Python版：レイアウト画像＋MarkdownからPPTX生成
├── config.json              # スライド内容・テキスト設定（make_pptx.py用）
├── sample_content.md        # Markdownサンプル（make_pptx_from_md.py用）
└── vba/
    └── make_pptx.bas        # VBAマクロ版
```

## 使い方①：config.json から生成（シンプル版）

```bash
cd ~/projects/ppt_auto
python3 make_pptx.py config.json
```

## 使い方②：レイアウト画像 + Markdown から生成（AI解析版）

```bash
# ANTHROPIC_API_KEY が必要
python3 make_pptx_from_md.py <レイアウト画像> <Markdownファイル> [出力.pptx]

# 例
python3 make_pptx_from_md.py layout.png sample_content.md
```

### Markdownフォーマット

```markdown
---
title: プレゼンタイトル
subtitle: サブタイトル
author: 作成者名
date: 2026年4月29日
---

## スライドタイトル

本文テキスト
- 箇条書き1
- 箇条書き2

## まとめ <!-- summary -->

まとめ本文
```

## 使い方（VBA版）

1. PowerPointを開き、Alt+F11でVBAエディタを起動
2. `vba/make_pptx.bas` をインポート
3. `MakePPTX` マクロを実行
