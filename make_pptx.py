"""
PowerPoint定型スライド自動生成スクリプト
使い方: python3 make_pptx.py config.json  （またはMarkdownファイルを直接指定: python3 make_pptx.py slides.md）
"""

import json
import re
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

# コンテンツ領域（タイトルバー直下から下部バー直上まで）
CONTENT_LEFT   = Inches(0.8)
CONTENT_WIDTH  = Inches(11.7)
CONTENT_TOP    = Inches(1.5)


# スライドサイズ（16:9）
SLIDE_WIDTH  = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

# カラー定義
COLOR_ACCENT  = RGBColor(0x1F, 0x49, 0x7D)  # 濃紺（タイトルバー）
COLOR_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK    = RGBColor(0x26, 0x26, 0x26)   # 本文テキスト
COLOR_LIGHT   = RGBColor(0xF2, 0xF2, 0xF2)  # 背景薄グレー


def set_bg(slide, color):
    """スライド背景色を設定する"""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, color):
    """塗りつぶし矩形を追加する"""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()  # 枠線なし
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=24, bold=False, color=COLOR_DARK,
                 align=PP_ALIGN.LEFT, wrap=True, autofit=False):
    """テキストボックスを追加する（\\n はPowerPointの段落として分割する。
    run.textに\\nをそのまま渡すとXML上は改行文字が入るだけでPowerPointは
    改行として描画しないため、行ごとに別段落を作る必要がある）。
    autofit=Trueの場合、ボックスに収まらない量のテキストが入ったときPowerPoint側で
    自動的にフォントサイズを縮小する（python-pptxはフォントメトリクスを持たないため
    事前計算はできず、PowerPointの「自動調整」機能に委ねる）"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    if autofit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "メイリオ"
    return txBox


def make_title_slide(prs, cfg):
    """表紙スライドを作成する"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白レイアウト
    set_bg(slide, COLOR_LIGHT)

    # 上部アクセントバー
    add_rect(slide, 0, 0, SLIDE_WIDTH, Inches(0.15), COLOR_ACCENT)

    # 中央タイトル背景（タイトルが\nで複数行になる場合に備えて高さを確保）
    add_rect(slide, Inches(1), Inches(1.8), Inches(11.33), Inches(2.7), COLOR_ACCENT)

    # タイトル
    add_text_box(
        slide, cfg["title"],
        Inches(1.3), Inches(2.0), Inches(10.8), Inches(1.6),
        font_size=40, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )

    # サブタイトル
    add_text_box(
        slide, cfg.get("subtitle", ""),
        Inches(1.3), Inches(3.6), Inches(10.8), Inches(0.8),
        font_size=24, color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )

    # 作成者・日付
    info = f"{cfg.get('date', '')}　{cfg.get('author', '')}"
    add_text_box(
        slide, info,
        Inches(1), Inches(5.5), Inches(11.33), Inches(0.5),
        font_size=16, color=COLOR_DARK, align=PP_ALIGN.RIGHT
    )

    # 下部アクセントバー
    add_rect(slide, 0, Inches(7.2), SLIDE_WIDTH, Inches(0.3), COLOR_ACCENT)


def make_toc_slide(prs, cfg):
    """目次スライドを作成する"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, COLOR_LIGHT)

    add_rect(slide, 0, 0, SLIDE_WIDTH, Inches(1.2), COLOR_ACCENT)
    add_text_box(
        slide, "目次",
        Inches(0.5), Inches(0.2), Inches(12), Inches(0.8),
        font_size=32, bold=True, color=COLOR_WHITE
    )

    contents = [s["title"] for s in cfg.get("slides", [])]
    body = "\n".join(f"  {i+1}.  {t}" for i, t in enumerate(contents))
    add_text_box(
        slide, body,
        Inches(1.0), Inches(1.5), Inches(11), Inches(5.5),
        font_size=22, color=COLOR_DARK, autofit=True
    )

    add_rect(slide, 0, Inches(7.2), SLIDE_WIDTH, Inches(0.3), COLOR_ACCENT)


def add_quote_box(slide, text, top):
    """引用（blockquote）を左アクセント線付きの帯として追加する"""
    height = Inches(0.9)
    add_rect(slide, CONTENT_LEFT, top, CONTENT_WIDTH, height, COLOR_LIGHT)
    add_rect(slide, CONTENT_LEFT, top, Inches(0.06), height, COLOR_ACCENT)
    add_text_box(
        slide, text,
        CONTENT_LEFT + Inches(0.25), top + Inches(0.08), CONTENT_WIDTH - Inches(0.4), height - Inches(0.16),
        font_size=16, color=COLOR_ACCENT, autofit=True
    )
    return top + height + Inches(0.15)


def add_images_row(slide, image_paths, top, max_height=Inches(4.2)):
    """画像を横並びで挿入する（アスペクト比は維持したまま、全画像の高さを揃える）"""
    n = len(image_paths)
    gap = Inches(0.2)
    img_width = Emu(int((CONTENT_WIDTH - gap * (n - 1)) / n))

    # 各画像のアスペクト比を先に調べ、列幅いっぱいに収めたときの高さの最小値を
    # 全画像共通の高さとする（アスペクト比がバラバラだと高さが不揃いになるため）
    aspects = []
    for path in image_paths:
        pic = slide.shapes.add_picture(path, 0, 0, width=img_width)
        aspects.append(pic.width / pic.height)
        pic._element.getparent().remove(pic._element)

    common_height = min(Emu(int(img_width / a)) for a in aspects)
    common_height = min(common_height, max_height)

    left = CONTENT_LEFT
    for path, aspect in zip(image_paths, aspects):
        width = Emu(int(common_height * aspect))
        pic_left = int(left + (img_width - width) / 2)
        slide.shapes.add_picture(path, pic_left, top, width=width, height=common_height)
        left += img_width + gap
    return top + common_height + Inches(0.15)


def make_content_slide(prs, slide_cfg):
    """本文スライドを作成する（本文・画像・引用の組み合わせに対応）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, COLOR_LIGHT)

    add_rect(slide, 0, 0, SLIDE_WIDTH, Inches(1.2), COLOR_ACCENT)
    add_text_box(
        slide, slide_cfg["title"],
        Inches(0.5), Inches(0.2), Inches(12), Inches(0.8),
        font_size=32, bold=True, color=COLOR_WHITE
    )

    body = slide_cfg.get("body", "")
    images = slide_cfg.get("images") or ([slide_cfg["image"]] if slide_cfg.get("image") else [])
    quote = slide_cfg.get("quote", "")

    cursor = CONTENT_TOP

    if body:
        # 画像がある場合は本文を短めに確保、ない場合は広く使う
        body_height = Inches(2.2) if images else Inches(4.6)
        add_text_box(
            slide, body,
            CONTENT_LEFT, cursor, CONTENT_WIDTH, body_height,
            font_size=20 if images else 22, color=COLOR_DARK, autofit=True
        )
        cursor += body_height

    if images:
        remaining = Inches(7.0) - cursor - (Inches(1.05) if quote else 0)
        cursor = add_images_row(slide, images, cursor, max_height=remaining)

    if quote:
        cursor = min(cursor, Inches(6.0))
        add_quote_box(slide, quote, cursor)

    add_rect(slide, 0, Inches(7.2), SLIDE_WIDTH, Inches(0.3), COLOR_ACCENT)


def make_summary_slide(prs, title, body):
    """まとめスライドを作成する（背景色を変えて強調）"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, COLOR_ACCENT)

    add_text_box(
        slide, title,
        Inches(0.5), Inches(0.3), Inches(12), Inches(0.9),
        font_size=36, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER
    )

    add_rect(slide, Inches(0.5), Inches(1.1), Inches(12.33), Inches(0.05), COLOR_WHITE)

    add_text_box(
        slide, body,
        Inches(0.8), Inches(1.4), Inches(11.7), Inches(5.5),
        font_size=22, color=COLOR_WHITE, autofit=True
    )


def _strip_inline_md(text: str) -> str:
    """行内の**太字**・*イタリック*・`コード`記法の記号だけを取り除く"""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text


def parse_markdown(md_path: Path) -> dict:
    """Markdownファイルをconfig.jsonと同じ構造のdictに変換する（config.json不要でMarkdownから直接生成するための入口）。

    対応フォーマット（Marpに似せたシンプルな記法）:
      ---
      title: タイトル
      subtitle: サブタイトル
      author: 作成者
      date: 日付
      ---

      ## スライドタイトル

      本文テキスト
      - 箇条書き（・付きに自動変換される）

      > 引用・強調したい一言（quoteフィールドになる）

      ![説明](image.png)（imagesフィールドになる。複数書けば横並び）

      ## まとめ <!-- summary -->

      まとめ本文

    フロントマター（---...---）は丸ごと省略可能。titleが指定されていない場合、
    本文の最初の"# "見出し（H1）、なければ最初の"## "見出し（H2）を自動でタイトルに
    使う。subtitle/author/dateは元々省略可（空欄）。これにより、README.md等の
    既存Markdownをそのまま渡しても動作する。
    """
    content = md_path.read_text(encoding="utf-8")
    cfg = {"title": "", "subtitle": "", "author": "", "date": "", "slides": []}

    # フロントマター（---...---）を解析（フロントマター自体が省略されていてもよい）
    fm_match = re.match(r"^---\s*\n([\s\S]*?)\n---\s*\n?", content)
    if fm_match:
        for line in fm_match.group(1).split("\n"):
            if ":" in line:
                k, _, v = line.partition(":")
                k, v = k.strip(), v.strip()
                if k in cfg:
                    cfg[k] = v
        content = content[fm_match.end():]

    # titleが未指定の場合、本文の見出しから自動取得する
    # （最初の"# "見出し、なければ最初の"## "見出しをタイトルとして使う。
    # subtitle/author/dateは元々省略可＝空欄でよいため、これでフロントマター自体が不要になる）
    if not cfg["title"]:
        m_h1 = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
        if m_h1:
            cfg["title"] = m_h1.group(1).strip()
        else:
            m_h2 = re.search(r"^##\s+(.+)$", content, re.MULTILINE)
            if m_h2:
                cfg["title"] = m_h2.group(1).strip()

    # H1/H2見出しでスライドに分割（H3以下はスライド内の小見出しとして扱う）
    parts = re.split(r"^#{1,2}\s+(.+)$", content, flags=re.MULTILINE)

    i = 1
    while i < len(parts) - 1:
        raw_title = parts[i].strip()
        raw_body = parts[i + 1].strip()

        slide_type = "content"
        if "<!-- summary -->" in raw_title:
            slide_type = "summary"
        raw_title = re.sub(r"<!--.*?-->", "", raw_title).strip()

        body_lines = []
        quote_lines = []
        image_paths = []
        in_code_block = False

        for line in raw_body.split("\n"):
            stripped = line.strip()

            # コードブロック（```...```）は省略する（スライドには不向きなため）
            if stripped.startswith("```"):
                in_code_block = not in_code_block
                continue
            if in_code_block:
                continue

            # 水平線（---, ___, ***）は無視する
            if re.fullmatch(r"[-_*]{3,}", stripped):
                continue

            m_img = re.match(r"!\[.*?\]\((.+?)\)", stripped)
            if m_img:
                image_paths.append(m_img.group(1))
                continue

            if stripped.startswith(">"):
                quote_lines.append(stripped.lstrip(">").strip())
                continue

            # H3〜H6はスライドを分けず、記号を外して小見出し行として扱う
            m_subhead = re.match(r"^#{3,6}\s+(.*)", stripped)
            if m_subhead:
                body_lines.append(_strip_inline_md(m_subhead.group(1)))
                continue

            m_bullet = re.match(r"^[-*+]\s+(.*)", stripped)
            prefix, text = ("・", m_bullet.group(1)) if m_bullet else ("", stripped)
            body_lines.append(prefix + _strip_inline_md(text))

        slide_cfg = {
            "type": slide_type,
            "title": raw_title,
            "body": "\n".join(body_lines).strip("\n"),
        }
        if image_paths:
            # 画像パスはMarkdownファイルからの相対パスとして解決する
            slide_cfg["images"] = [
                str((md_path.parent / p).resolve()) if not Path(p).is_absolute() else p
                for p in image_paths
            ]
        if quote_lines:
            slide_cfg["quote"] = _strip_inline_md(" ".join(quote_lines))

        cfg["slides"].append(slide_cfg)
        i += 2

    return cfg


MARKDOWN_TEMPLATE_HELP = """\
使い方: python3 make_pptx.py config.json
       python3 make_pptx.py slides.md

Markdownファイルの記法:

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

補足: フロントマター（---...---）は省略可能。titleを省略した場合、本文の
最初の "# " 見出し（なければ最初の "## " 見出し）を自動でタイトルに使う。
subtitle/author/dateも省略時は空欄になるだけなので、既存のMarkdownファイル
（README.md等）をそのまま渡しても動作する。
"""


def main():
    if len(sys.argv) < 2:
        print("使い方: python3 make_pptx.py config.json")
        print("       python3 make_pptx.py slides.md")
        sys.exit(1)

    if sys.argv[1] == "?":
        print(MARKDOWN_TEMPLATE_HELP)
        sys.exit(0)

    config_path = Path(sys.argv[1])
    if not config_path.exists():
        print(f"エラー: {config_path} が見つかりません")
        sys.exit(1)

    if config_path.suffix.lower() == ".md":
        cfg = parse_markdown(config_path)
    else:
        with open(config_path, encoding="utf-8") as f:
            cfg = json.load(f)

    prs = Presentation()
    prs.slide_width  = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # 表紙
    make_title_slide(prs, cfg)

    # 目次（スライドが2枚以上ある場合のみ）
    if len(cfg.get("slides", [])) >= 2:
        make_toc_slide(prs, cfg)

    # 本文・まとめ
    for slide_cfg in cfg.get("slides", []):
        if slide_cfg.get("type") == "summary":
            make_summary_slide(prs, slide_cfg["title"], slide_cfg["body"])
        else:
            make_content_slide(prs, slide_cfg)

    out_path = config_path.with_name(config_path.stem + ".pptx")
    prs.save(out_path)
    print(f"生成完了: {out_path}")


if __name__ == "__main__":
    main()
