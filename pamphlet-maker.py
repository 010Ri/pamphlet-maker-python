import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt, Cm
# from pptx.enum.text import PP_ALIGN
# from pptx.enum.text import MSO_ANCHOR
import datetime

now = datetime.datetime.now()  # 現在時刻の取得
today = now.strftime('%Y年%m月%d日%H時%M分%S秒')  # 現在時刻を年月曜日で表示

save_name = './files/' + today  # 保存用のパワポのファイル名
pt = Presentation()  # Presentationクラスをインスタンス化

# JIS B5 縦向き にスライドサイズを変更
pt.slide_height=Cm(25.7)
pt.slide_width=Cm(18.2)

# 白紙のページを追加
slide_layout = pt.slide_layouts[6]
slide = pt.slides.add_slide(slide_layout)

# テキストボックスを挿入
textbox = slide.shapes.add_textbox(0, 0, Cm(15), Cm(5))  # (x座標, y座標, 横幅, 縦幅)
text_frame = textbox.text_frame
# text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
# text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
text_frame.text = 'テキストボックスを挿入しました。'

# テキストボックスに段落の追加
p = text_frame.add_paragraph()
p.text = 'テキストボックスに段落を追加しました。'

# 白紙のページを追加
slide_layout = pt.slide_layouts[6]
slide = pt.slides.add_slide(slide_layout)

# 画像の挿入
pic = slide.shapes.add_picture('/Users/kosei/Desktop/layout-flame/layout-flame-png/レイアウト枠_1L.png', 0, 0, Cm(18.2), Cm(25.7))  # (file path, x座標, y座標, 横幅, 縦幅)

# 画像の回転
# pic.rotation = 20


"""
# 1枚目
title_slide_layout = pt.slide_layouts[0]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "Pythonでパワポを作成してみた"
subtitle.text = today

# 2枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "目次"
body_shape = slide.placeholders[1]

p = body_shape.text_frame.add_paragraph()
p.text = '1. pythonでパワポを作るために'
p.level = 1
p = body_shape.text_frame.add_paragraph()
p.text = '2. 必要なライブラリのインストール'
p.level = 1
p = body_shape.text_frame.add_paragraph()
p.text = '3. 動作確認'
p.level = 1
p = body_shape.text_frame.add_paragraph()
p.text = '4. まとめ'
p.level = 1

# 3枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
title.text = "pythonでパワポを作るために"

body_shape = slide.placeholders[1]
p = body_shape.text_frame.add_paragraph()
p.text = '1. python-pptxを使えば実現可能！'
p.level = 1

# 4枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "2. 必要なライブラリのインストール"
body_shape = slide.placeholders[1]
p = body_shape.text_frame.add_paragraph()
p.text = '次のコマンドをターミナルにて実行'
# p.level = 2
p = body_shape.text_frame.add_paragraph()
p.text = 'pip install python-pptx'
p.level = 1

# 5枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
title.text = "3. 動作確認"
body_shape = slide.placeholders[1]

p = body_shape.text_frame.add_paragraph()
p.text = "プログラムを動かしてみよう！"
p.level = 1
p.font.bold = True

# 6枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "4. まとめ"
body_shape = slide.placeholders[1]
subtitle.text = "python-pptxを使えば、簡単にパワポが作れちゃう！！"

# 7枚目
title_slide_layout = pt.slide_layouts[1]
slide = pt.slides.add_slide(title_slide_layout)
title = slide.shapes.title
title.text = "7. 画像を追加する"
# 画像をスライドに追加
img_path = "./img/sample.jpg"
slide.shapes.add_picture(img_path, Inches(1.5), Inches(2))

"""

pt.save(save_name)
print("ファイルの書き出し完了しました")