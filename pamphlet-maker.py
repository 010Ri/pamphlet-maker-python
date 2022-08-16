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

# layout-flame_1-L , layout-flame_1(8)-Lのテキストボックス調整用の関数
def modify_textbox_position_1_L(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(12.35)
    textbox.left = Cm(-2.73)

# layout-flame_1-R , layout-flame_1(8)-Rのテキストボックス調整用の関数
def modify_textbox_position_1_R(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(12.35)
    textbox.left = Cm(10.95)

# layout-flame_2-2-Lのテキストボックス調整用の関数
def modify_textbox_position_2_2_L(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(-2.7)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(18.195)
    textbox.left = Cm(-2.7)

# layout-flame_2-2-Rのテキストボックス調整用の関数
def modify_textbox_position_2_2_R(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(10.95)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(18.195)
    textbox.left = Cm(10.95)

# layout-flame_2-4-4-L , layout-flame_2-4-8-Lのテキストボックス調整用の関数
def modify_textbox_position_2_4_L(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(-2.65)
    # テキストボックスを挿入（中側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(15.3)
    textbox.left = Cm(-0.165)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(21.1)
    textbox.left = Cm(-0.165)

# layout-flame_2-4-4-R , layout-flame_2-4-8-Rのテキストボックス調整用の関数
def modify_textbox_position_2_4_R(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(6.535)
    textbox.left = Cm(10.95)
    # テキストボックスを挿入（中側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(15.12)
    textbox.left = Cm(13.45)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(20.92)
    textbox.left = Cm(13.45)

# layout-flame_3-4-L , layout-flame_3-8-L , layout-flame_3(4)-4-L , layout-flame_3(4)-8-L , layout-flame_3(8)-4-Lのテキストボックス調整用の関数
def modify_textbox_position_3_L(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(9.55)
    textbox.left = Cm(-2.7)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(21.1)
    textbox.left = Cm(-0.165)

# layout-flame_3-4-R , layout-flame_3-8-R , layout-flame_3(4)-4-R , layout-flame_3(4)-8-R , layout-flame_3(8)-8-Rのテキストボックス調整用の関数
def modify_textbox_position_3_R(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(10), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告(会社名)・イベント名'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(9.4)
    textbox.left = Cm(10.95)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(21)
    textbox.left = Cm(13.45)

# layout-flame_4-4-4-4-L , layout-flame_4-4-4-8-Lのテキストボックス調整用の関数
def modify_textbox_position_4_L(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(3.8)
    textbox.left = Cm(0.1)
    # テキストボックスを挿入（中上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(9.567)
    textbox.left = Cm(0.1)
    # テキストボックスを挿入（中下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(15.433)
    textbox.left = Cm(0.1)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = -90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(21.2)
    textbox.left = Cm(0.1)

# layout-flame_4-4-4-4-R , layout-flame_4-4-4-8-Rのテキストボックス調整用の関数
def modify_textbox_position_4_R(id):
    slide = pt.slides[id]  # layout-flame_something
    # テキストボックスを挿入（上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(3.7)
    textbox.left = Cm(13.08)
    # テキストボックスを挿入（中上側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(9.52)
    textbox.left = Cm(13.08)
    # テキストボックスを挿入（中下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(15.25)
    textbox.left = Cm(13.08)
    # テキストボックスを挿入（下側）
    textbox = slide.shapes.add_textbox(0, 0, Cm(5), Cm(1))  # (x座標, y座標, 横幅, 縦幅)
    textbox.text = '広告'
    textbox.rotation = 90
    text_frame = textbox.text_frame
    text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.autosize = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    textbox.top = Cm(21.05)
    textbox.left = Cm(13.08)
    

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
text_frame.text = 'python-pptxでパワポを自動生成するやつ'
# テキストボックスに段落の追加
p = text_frame.add_paragraph()
p.text = 'テキストボックスの位置は適当です'


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
    elif file_name == "layout-flame_2-4-4-L" : 
        id = 7  # layout-flame_2-4-4-L
        modify_textbox_position_2_4_L(id)
    elif file_name == "layout-flame_2-4-4-R" : 
        id = 8  # layout-flame_2-4-4-R
        modify_textbox_position_2_4_R(id)
    elif file_name == "layout-flame_2-4-8-L" : 
        id = 9  # layout-flame_2-4-8-L
        modify_textbox_position_2_4_L(id)
    elif file_name == "layout-flame_2-4-8-R" : 
        id = 10  # layout-flame_2-4-8-R
        modify_textbox_position_2_4_R(id)
    elif file_name == "layout-flame_3-4-L" : 
        id = 11  # layout-flame_3-4-L
        modify_textbox_position_3_L(id)
    elif file_name == "layout-flame_3-4-R" : 
        id = 12  # layout-flame_3-4-R
        modify_textbox_position_3_R(id)
    elif file_name == "layout-flame_3-8-L" : 
        id = 13  # layout-flame_3-8-L
        modify_textbox_position_3_L(id)
    elif file_name == "layout-flame_3-8-R" : 
        id = 14  # layout-flame_3-8-R
        modify_textbox_position_3_R(id)
    elif file_name == "layout-flame_3(4)-4-L" : 
        id = 15  # layout-flame_3(4)-4-L
        modify_textbox_position_3_L(id)
    elif file_name == "layout-flame_3(4)-4-R" : 
        id = 16  # layout-flame_3(4)-4-R
        modify_textbox_position_3_R(id)
    elif file_name == "layout-flame_3(4)-8-L" : 
        id = 17  # layout-flame_3(4)-8-L
        modify_textbox_position_3_L(id)
    elif file_name == "layout-flame_3(4)-8-R" : 
        id = 18  # layout-flame_3(4)-8-R
        modify_textbox_position_3_R(id)
    elif file_name == "layout-flame_3(8)-4-L" : 
        id = 19  # layout-flame_3(8)-4-L
        modify_textbox_position_3_L(id)
    elif file_name == "layout-flame_3(8)-8-R" : 
        id = 20  # layout-flame_3(8)-8-R
        modify_textbox_position_3_R(id)
    elif file_name == "layout-flame_4-4-4-4-L" : 
        id = 21  # layout-flame_4-4-4-4-L
        modify_textbox_position_4_L(id)
    elif file_name == "layout-flame_4-4-4-4-R" : 
        id = 22  # layout-flame_4-4-4-4-R
        modify_textbox_position_4_R(id)
    elif file_name == "layout-flame_4-4-4-8-L" : 
        id = 23  # layout-flame_4-4-4-8-L
        modify_textbox_position_4_L(id)
    elif file_name == "layout-flame_4-4-4-8-R" : 
        id = 24  # layout-flame_4-4-4-8-R
        modify_textbox_position_4_R(id)


# ファイルを任意の名前で保存（現在時刻をファイル名として保存するようにしている）
now = datetime.datetime.now()  # 現在時刻の取得
today = now.strftime('%Y年%m月%d日%H時%M分%S秒')  # 現在時刻を年月曜日で表示
save_name = '/Users/kosei/Desktop/output-files/' + today  # 保存用のパワポのファイル名
pt.save(save_name)
print("ファイルの書き出し完了しました")
