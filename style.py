# -*- coding: utf-8 -*-

import sys
import  xlsxFormatSetting as settings

reload(sys)
sys.setdefaultencoding('utf8')
reload(sys)

class Style:
    #Style的构造函数 Style(*style) 为不定参函数
    #当构造函数无参时，使用默认的样式
    #当构造函数参数有一个参数时，采用xlsxFormatSetting中style中对应的样式，无定义部分采用默认样式
    #当构造函数参数出现其他情况，采用默认的样式
    def __init__(self, *style):
        #默认字体为Arial，8，不加粗，黑色
        #默认背景色为空，默认颜色为白色
        #默认文本样式为水平居中，垂直居中，自动换行
        #默认边框为细线，黑色
        #字体格式
        font_family = 'Arial'
        font_size = 8
        font_blod = False
        font_color = 'FF000000'
        patternFill_patternType = None
        patternFill_color = 'FFFFFFFF'
        alignment_horizontal = 'center'
        alignment_vertical = 'center'
        alignment_wrap_text = True
        border_style = 'thin'
        border_color = 'FF000000'
        number_format_number = '#,##0'

        if len(style) == 1 and settings.style.has_key(style[0]):
            style = settings.style[style[0]]
        else:
            style = {}
        self.font = {'family': style['font_family'] if style.has_key('font_family') else font_family,
                     'size': style['font_size'] if style.has_key('font_size') else font_size,
                     'blod': style['font_blod'] if style.has_key('font_blod') else font_blod,
                     'color': style['font_color'] if style.has_key('font_color') else font_color}
        self.patternFill = {'patternType': 'solid' \
                                if style.has_key('patternFill_color') else patternFill_patternType,
                            'color': style['patternFill_color'] \
                                if style.has_key('patternFill_color') else patternFill_color}
        self.alignment = {'horizontal': style['alignment_horizontal'] \
                                if style.has_key('alignment_horizontal') else alignment_horizontal,
                          'vertical': style['alignment_vertical'] \
                                if style.has_key('alignment_vertical') else alignment_vertical,
                          'wrap_text': style['alignment_wrap_text'] \
                                if style.has_key('alignment_wrap_text') else alignment_wrap_text}
        self.border = {'style': 'thin' if style.has_key('border_color') else border_style,
                       'color': style['border_color'] if style.has_key('border_color') else border_color}
        self.number_format = {'number': style['number_format_number'] \
                                if style.has_key('number_format_number') else number_format_number}

    def setFont(self, name, size, blod, color):
        self.font['name'] = name
        self.font['size'] = size
        self.font['blod'] = blod
        self.font['color'] = 'FF' + color

    def setPatternFill(self, color):
        self.patternFill['patternType'] = 'solid'
        self.patternFill['color'] = 'FF' + color

    def setborder(self, color):
        self.border['style'] = 'thin'
        self.border['color'] = 'FF' + color

if __name__ == '__main__':
    style = Style('标题')
    print style.font
    print style.patternFill
    style = Style()
    print style.font
    print style.patternFill