from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# プレゼンテーションの作成
prs = Presentation()

# --- スライド1: タイトル ---
slide_layout = prs.slide_layouts[0]  # タイトルスライド
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

# タイトルのフォント設定
title.text = "XX技術の進化とその応用"
title_format = title.text_frame.paragraphs[0].font
title_format.size = Pt(40)
title_format.bold = True
title_format.color.rgb = RGBColor(0, 51, 102)  # 青色

# サブタイトルのフォント設定
subtitle.text = "山田 太郎\n2025年3月20日"
subtitle_format = subtitle.text_frame.paragraphs[0].font
subtitle_format.size = Pt(18)
subtitle_format.color.rgb = RGBColor(102, 102, 102)  # グレー

# 背景色を変更
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(240, 240, 240)  # 背景を淡いグレーに

# --- スライド2: 研究の目的（背景） ---
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
title.text = "研究の目的"
content = slide.shapes.placeholders[1]
content.text = ("・XX技術の現状と課題\n"
                "・技術の進化に向けたニーズ\n"
                "・本研究の目的は、XX技術の応用を拡大すること")

# タイトルフォント変更
title_format = title.text_frame.paragraphs[0].font
title_format.size = Pt(28)
title_format.bold = True
title_format.color.rgb = RGBColor(0, 51, 102)

# 背景色の変更
background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白色背景

# --- スライド3: 画像の追加 ---
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
title.text = "研究の動機・課題設定"

# 画像をスライドに追加
img_path = "path_to_image.png"  # 画像ファイルのパスを指定
# サイズ調整：width=4インチ、高さ=3インチ
slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), width=Inches(6), height=Inches(3.5))

# --- スライド4: グラフや図の挿入（任意） ---
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
title.text = "研究結果"

# グラフや図を追加したい場合は、`matplotlib`などを使ってグラフを描画し、画像として保存して挿入できます

# --- スライド5: 結論 ---
slide_layout = prs.slide_layouts[1]
slide = prs.slides.add_slide(slide_layout)
title = slide.shapes.title
title.text = "結論"
content = slide.shapes.placeholders[1]
content.text = ("・本研究の要点\n"
                "・XX技術の重要性\n"
                "・研究の意義と社会への貢献")

# タイトルフォント変更
title_format = title.text_frame.paragraphs[0].font
title_format.size = Pt(28)
title_format.bold = True
title_format.color.rgb = RGBColor(0, 51, 102)

# プレゼンテーションの保存
prs.save('styled_research_presentation.pptx')

print("シンプルで見やすいデザインのPowerPointファイル 'styled_research_presentation.pptx' が作成されました。")
