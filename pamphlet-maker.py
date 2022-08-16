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

# layout-flame_1-Lとlayout-flame_1(8)-Lのテキストボックス調整用の関数
def modify_textbox_position_1_L(id):
    # 1-Lのテキストボックス位置調整
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(12.35)
    textbox.left = Cm(-2.73)

# layout-flame_1-Rとlayout-flame_1(8)-Rのテキストボックス調整用の関数
def modify_textbox_position_1_R(id):
    # 1-Lのテキストボックス位置調整
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(12.35)
    textbox.left = Cm(10.95)

# layout-flame_2-2-Lのテキストボックス調整用の関数
def modify_textbox_position_2_2_L(id):
    # 1-Lのテキストボックス位置調整
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(-2.73)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(18.195)
    textbox.left = Cm(-2.73)

# layout-flame_2-2-Rのテキストボックス調整用の関数
def modify_textbox_position_2_2_R(id):
    # 1-Lのテキストボックス位置調整
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(10.95)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = 'ここに会社名が入ります'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(18.195)
    textbox.left = Cm(10.95)


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


# ファイル名リストからファイル名を順番に取り出す
for file_name in file_name_list:
    # 該当するレイアウト枠を挿入した新規ページを追加
    add_page(file_name)

    # テキストボックス調整
    if file_name == "layout-flame_1-L" :
        id = 1  # layout-flame_1-L
        modify_textbox_position_1_L(id)
    elif file_name == "layout-flame_1-R" :
        id = 2  # layout-flame_1-R
        modify_textbox_position_1_R(id)
    elif file_name == "layout-flame_1(8)-L" : 
        id = 3  # layout-flame_1(8)-L
        modify_textbox_position_1_L(id)
    elif file_name == "layout-flame_1(8)-R" : 
        id = 4  # layout-flame_1(8)-R
        modify_textbox_position_1_R(id)
    elif file_name == "layout-flame_2-2-L" : 
        id = 5  # layout-flame_2-2-L
        modify_textbox_position_2_2_L(id)
    elif file_name == "layout-flame_2-2-R" : 
        id = 6  # layout-flame_2-2-R
        modify_textbox_position_2_2_R(id)


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
