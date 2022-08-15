import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt, Cm, Mm
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
import datetime

# Presentationクラスをインスタンス化
pt = Presentation()

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

# 白紙のページを追加した後、レイアウト枠の画像を挿入するための関数
def add_page(file_name):
    # 白紙のページを追加
    slide_layout = pt.slide_layouts[6]
    slide = pt.slides.add_slide(slide_layout)
    # 画像の挿入
    file_path = "./img/" + file_name + ".png"
    pic = slide.shapes.add_picture(file_path, 0, 0, Cm(18.2), Cm(25.7))  # (file path, x座標, y座標, 横幅, 縦幅)
    # 画像の回転
    # pic.rotation = 20


# file name list
file_name_list = [
    "layout-flame_1-L",
    "layout-flame_1-R",
    "layout-flame_1(8)-L",
    "layout-flame_1(8)-R",
    "layout-flame_2-2-L",
    "layout-flame_2-2-R",
    "layout-flame_2-4-4-L",
    "layout-flame_2-4-4-R",
    "layout-flame_2-4-8-L",
    "layout-flame_2-4-8-R",
    "layout-flame_3-4-L",
    "layout-flame_3-4-R",
    "layout-flame_3-8-L",
    "layout-flame_3-8-R",
    "layout-flame_3(4)-4-L",
    "layout-flame_3(4)-4-R",
    "layout-flame_3(4)-8-L",
    "layout-flame_3(4)-8-R",
    "layout-flame_3(8)-4-L",
    "layout-flame_3(8)-8-R",
    "layout-flame_4-4-4-4-L",
    "layout-flame_4-4-4-4-R",
    "layout-flame_4-4-4-8-L",
    "layout-flame_4-4-4-8-R"
]
# print(len(file_list))


# ファイル名リストからファイル名を順番に取り出して、該当するレイアウト枠を挿入した新規ページを追加
for file_name in file_name_list:
    add_page(file_name)

# 1-Lのテキストボックス
slide = pt.slides[1]  # 2ページ目
# テキストボックスを挿入
textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
textbox.text = 'ここに会社名が入ります'
textbox.rotation = -90
text_frame = textbox.text_frame
text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
textbox.top = Cm(12.35)
textbox.left = Cm(-2.7)

"""
# 図形の挿入
rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Cm(5), Cm(3)) 
# 図形のスタイル設定
rect.fill.solid() #塗りつぶし
rect.fill.fore_color.rgb = RGBColor(0, 0, 255) #塗りつぶし色の指定
rect.line.width = Mm(2) #枠線の指定、今回は2mm
rect.line.color.rgb = RGBColor(255, 0, 0) #枠線の色指定
# 図形にテキストを挿入
pg = rect.text_frame.paragraphs[0] # 図形からtext_frameオブジェクトを取り出し、1つ目のパラグラフを取得
pg.text = '図形です。'
pg.font.size = Pt(20) # テキストサイズ、今回は20ポイント
"""

# ファイルを任意の名前で保存（現在時刻をファイル名として保存するようにしている）
now = datetime.datetime.now()  # 現在時刻の取得
today = now.strftime('%Y年%m月%d日%H時%M分%S秒')  # 現在時刻を年月曜日で表示
save_name = '/Users/kosei/Desktop/output-files/' + today  # 保存用のパワポのファイル名
pt.save(save_name)
print("ファイルの書き出し完了しました")
