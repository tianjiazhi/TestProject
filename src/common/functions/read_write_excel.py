# coding = utf-8
"""
@Time      : 2019/10/19 0019 22:08
@Author    : YunFan
@File      : read_write_excel.py
@Software  : PyCharm
@Version   : 1.0
@Description: 
"""
import xlwt
import xlrd
import os
from src.common.functions.log import logger
from xlutils.copy import copy


log = logger

class Reader():
    """用来读取excel文件中的内容"""
    def __init__(self):
        self.workbook = None   # 待要读取的工作薄
        self.sheet = None      # 待要读取的sheet页
        self.rows = 0          # 初始化当前sheet页的总行数
        self.r = 0             # 初始化逐行读取时的行数

    def open_excel(self,srcfile):

        if not os.path.isfile(srcfile):  # 判断要读取的excel文件是否存在
            log.info("Error:{}文件不存在!".format(srcfile))
            return

        # 设置读取excel使用的utf8编码
        xlrd.Book.encoding = "utf8"
        # 读取excel内容到缓存workbook
        self.workbook = xlrd.open_workbook(filename=srcfile)
        # 默认读取第一个sheet页面
        self.sheet = self.workbook.sheet_by_index(0)
        # 设置rows为当前sheet页的总行数
        self.rows = self.sheet.nrows
        # 读取时从第一行读取
        self.r = 0
        return

    def get_sheets(self):
        '''获取sheet页面名称，并返回一个列表。'''

        sheets = self.workbook.sheet_names()
        log.info("读取的excel文件中包含以下sheet页面：{}".format(sheets))
        return sheets

    def set_sheet(self,name):
        '''通过sheet页面名称，切换sheet页'''
        self.sheet = self.workbook.sheet_by_name(name)
        self.rows = self.sheet.nrows
        self.r = 0
        return

    def readline(self):
        '''逐行读取指定sheet页面的数据'''
        row = None
        # 如果当前还没有读到最后一行，则读取下一行

        if (self.r < self.rows):
            row = self.sheet.row_values(self.r)
            self.r = self.r + 1

        return row

class CopyFileWriter():
    """复制已有文件并在该文件中写入数据"""
    def __init__(self):
        self.workbook = None  # 读取需要复制的excel
        self.wb = None        # 拷贝的工作空间
        self.sheet = None     # 当前工作的sheet页
        self.df = None        # 记录生成的文件，用来保存
        self.row = 0          # 记录写入的行
        self.clo = 0          # 记录写入的列

    def copy_open_workbook(self,srcfile,dstfile):
        """
        复制并打开excel
        :param srcfile: 被复制的文件路径和名称
        :param dstfile: 复制后的文件路径和名称
        :return:
        """
        if not os.path.isfile(srcfile):            # 判断要复制的文件是否存在
            log.info("Error:{}文件不存在!".format(srcfile))

        if os.path.isfile(dstfile):                # 判断要新建的文档是否存在，存在则提示。
            log.info("Warning：" + dstfile + "文件已经存在!")

        self.df = dstfile                            # 记录要保存的文件
        # 获取excel到缓存 ,其中：formatting_info是带格式的复制
        self.workbook = xlrd.open_workbook(filename = srcfile,formatting_info = True)
        self.wb = copy(self.workbook)                # 复制打开的excel文件
        # self.sheet = self.wb.get_sheet('Sheet1')   # 默认使用第一个sheet
        self.sheet = self.get_sheets()[0]  # 默认使用第一个sheet

        return

    def get_sheets(self):
        '''获取sheet页面名称，并返回一个列表。'''
        sheets = self.workbook.sheet_names()
        log.info("打开的excel文件中包含以下sheet页面：{}".format(sheets))
        return sheets

    def set_sheet(self,name):
        '''通过sheet名称，切换sheet页面'''
        self.sheet = self.wb.get_sheet(name)
        return

    def write(self,r,c,value):
        """写入指定的单元格，保留原格式"""

        # 获取要写入的单元格
        def _getCell(sheet,r,c):

            # 获取行
            row = sheet._Worksheet__rows.get(r)
            if not row:
                return None

            # 获取单元格
            cell = row._Row__cells.get(c)
            return cell

        # 获取要写入的单元格
        cell = _getCell(self.sheet,r,c)

        # 向获取的单元格中写值
        self.sheet.write(r,c,value)

        # 单元格格式在写入前后保持一致。
        if cell:      # 获取要写入的单元格
            ncell = _getCell(self.sheet,r,c)
            if ncell: # 设置写入后格式和写入前一致
                ncell.xf_idx = cell.xf_idx

        return

    def save_closed(self):
        '''保存并关闭文件'''
        self.wb.save(self.df)
        return

class SetCellsStyle():

    def __init__(self):
        self.style = xlwt.XFStyle()  # 初始化样式

    def set_cells_colour(self,fore_colour = None):
        '''设置单元格背景颜色'''
        pattern = xlwt.Pattern()                     # 创建一个模式
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # 设置背景颜色的模式
        pattern.pattern_fore_colour = fore_colour    # 设置背景颜色
        #  设置单元格背景颜色 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta,
        self.style.pattern = pattern
        return self.style

    def set_font_style(self,font_name,font_colour,font_bold,font_height,font_italic,font_underline):
        """设置字体样式"""
        font = xlwt.Font()               # 为样式创建字体对象
        font.name = font_name            # 字体名称
        font.colour_index = font_colour  # 字体颜色      默认 黑色
        font.bold = font_bold            # 字体是否加粗   默认 不加组
        font.height = font_height        # 字体大小
        font.italic = font_italic        # 是否设置为斜体字
        font.underline = font_underline  # 下划线
        self.style.font = font           # 将赋值好的模式参数导入Style
        return self.style

    def set_cells_borders(self,left_borders,right_borders,top_borders,bottom_borders,borders_colour):
        '''设置单元格边框样式'''
        borders = xlwt.Borders()
        """
            * 细实线:1，    小粗实线:2，   细虚线:3，   中细虚线:4，    大粗实线:5，   双线:6，   细点虚线:7
            * 大粗虚线:8，   细点划线:9，  粗点划线:10， 细双点划线:11， 粗双点划线:12， 斜点划线:13
        """
        borders.left = left_borders             # 左边边框
        borders.right = right_borders           # 右边边框
        borders.top = top_borders               # 顶部边框
        borders.bottom = bottom_borders         # 底部边框

        borders.left_colour = borders_colour    # 左边框颜色
        borders.right_colour = borders_colour   # 右边框颜色
        borders.top_colour = borders_colour     # 上边框颜色
        borders.bottom_colour = borders_colour  # 下边框颜色

        self.style.borders = borders           # 将赋值好的模式参数导入Style
        return self.style

    def set_cells_alignment(self,horz,vert,wrap):
        '''设置单元格对齐方式'''
        alignment = xlwt.Alignment()
        alignment.horz = horz   # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.vert = vert   # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
        alignment.wrap = wrap      # 设置自动换行(0->不换行，1换行)
        self.style.alignment = alignment
        return self.style

class NewFileWriter():
    def __init__(self):
        self.workbook = None    # 新建的excel文件
        self.sheet = None       # 当前新建的sheet页
        self.row = 0            # 记录写入的行
        self.clo = 0            # 记录写入的列
        self.write_text = None  # 单元格中写入的内容
        # self.set_cells_style = SetCellsStyle()

    def new_workbook(self):  # 新建工作簿
        self.workbook = xlwt.Workbook(encoding='utf-8')

    def new_sheet(self, sheet_name): # 新建工作表
        self.sheet = self.workbook.add_sheet(sheet_name)  # 新建sheet

    def set_col_width(self,col:int,width:int=10):
        """
        设置当前工作表的 列宽
        :param clo:  列数
        :param width: 设置宽度（字符个数），
        """
        cols = self.sheet.col(col)
        cols.width = 256 * width   # 256位衡量单位，width为字符数

    def set_row_width(self,row:int,font_height:int=240,font_name:str='Times New Roman',font_color:str='black',boldface:str='off'):
        """
        设置当前工作表的 行高
        :param row: 行数
        :param font_height:设置第row行的字体大小  240像素 = 20 * 字号
        :param font_name:  设置第row行的字体名称
        :param font_color: 设置第row行的字体颜色
        :param boldface:   设置第row行的字体是否加黑
        :return:
        """

        # 'font:height 720,name Times New Roman,color-index 20,bold on'
        tall_style_str =  'font:height {},name {},color-index {},bold {}'.format(str(font_height),font_name,font_color,boldface)

        tall_style = xlwt.easyxf(tall_style_str)  # 12pt,类型小初的字号
        rows = self.sheet.row(row)
        rows.set_style(tall_style)

    def write_data_cell(self,rows,columns,write_texts): # 向单元格中写入数据
        self.row = rows      # 行数
        self.clo = columns   # 列数
        self.write_text = write_texts
        self.sheet.write(self.row,self.clo,self.write_text) #  ,self.set_cells_style)

    def merge_cells(self,rows,across_row,columns,across_clo,write_texts):# 合并单元格并写入数据
        self.row = rows
        self.clo = columns
        current_row = self.row + across_row
        current_clo = self.clo + across_clo
        self.write_text = write_texts
        self.sheet.write_merge(self.row,current_row,self.clo,current_clo,self.write_text)# ,self.set_cells_style)

    def save_workbook(self, save_full_file_path): # # 对文件进行保存
        if os.path.isfile(save_full_file_path):
            # 判断要新建的文档是否存在，存在则提示。
            log.info("Warning：" + save_full_file_path + "文件已经存在!")
        self.workbook.save(save_full_file_path)


# 测试代码
if __name__ == "__main__":
    pass
    #
    # srcfile = "..\\files\\Template\\Interface_TD_Template.xls"  # 读取的excel文件
    # dstfile = "..\\files\\TestCase\\InterfaceTestDate.xls"  # 读取后保存的excel文件
    #
    # # 读取一个excel工作表中所有的sheet页中的所有行
    # reader = Reader()
    # reader.open_excel(srcfile)
    # sheetName = reader.get_sheets()
    # for sheet in sheetName:
    #     reader.set_sheet(sheet)
    #     for i in range(reader.rows):
    #         log.info(reader.readline())
    #
    #
    # # 向excel工作簿的指定的sheet页面写入数值
    # writer = CopyFileWriter()
    # writer.copy_open_workbook(srcfile,dstfile)
    # sheetName = writer.get_sheets()
    # writer.set_sheet(sheetName[0])
    # writer.write(1, 12, "yunfan")
    # writer.save_closed()
    #
    # # 测试NewFileWriter
    # a = NewFileWriter()
    # a.new_workbook()
    # a.new_sheet("sheet4")
    #
    # a.set_col_width(0)
    # a.set_row_width(0)
    #
    # a.write_data(0, 0, u"张三张三张三张三张三张三张三张三张三张三张三张三")
    # a.write_data(1, 0, u"张三张三张三张三张三张三张三张三张三张三张三张三")
    # a.save_workbook("C:\\Users\\Administrator\\Desktop\\test_file.xls")















































































































































































# def set_cells_style_port(self, **kwargs):
#     # 默认参数
#     dict_data = {'fore_colour': 4,  # 单元格默认背景颜色值
#                  'font_name': 'Times New Roman',  # 默认字体名称
#                  'font_colour': 0,  # 字体颜色
#                  'font_bold': False,  # 字体是否加粗
#                  'font_height': 7,  # 默认字体大小
#                  'left_borders_size': 5,  # 左边边框
#                  'right_borders_size': 5,  # 右边边框
#                  'top_borders_size': 5,  # 顶部边框
#                  'bottom_borders_size': 5  # 底部边框
#                  }
#     # 获取默认数据中所有的健
#     dict_data_keys_list = list(dict_data.keys())
#     param_data = kwargs  # 用户传入的字典类型的参数对象
#
#     for key, value in param_data.items():
#         # 判处系统中未定义的字段
#         if key not in dict_data_keys_list:
#             return "参数:%s 系统为定义！，请检查后重试！" % key
#
#         # 判断：fore_colour字段的值-》背景颜色
#         if key == 'fore_colour' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型，请检查后重试！" % key
#         if key == 'fore_colour' and (value < 0x00 or value > 0xff):
#             return "参数:“%s” 的值不在0~255之间，请检查后重试！" % key
#
#         # 判断：font_name字段的值-》字体名称
#         if key == 'font_name' and not isinstance(value, str):
#             return "参数:“%s” 的值不是字符串类型，请检查后重试！" % key
#         if key == 'font_name' and len(value) == 0:
#             return "参数:“%s” 的值不可为空，请检查后重试！" % key
#
#         # 判断：font_colour字段的值-》字体颜色
#         if key == 'font_colour' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型，请检查后重试！" % key
#         if key == 'font_colour' and (value < 0x00 or value > 0xff):
#             return "参数:“%s” 的值不在0~255之间，请检查后重试！" % key
#
#         # 判断：font_bold字段的值-》字体是否加粗
#         if key == 'font_bold' and not isinstance(value, bool):
#             return "参数:“%s” 的值不是布尔类型的数据，请检查后重试！" % key
#
#         # 判断：font_height字段的值-》字体大小
#         if key == 'font_height' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型的数据，请检查后重试！" % key
#         if key == 'font_height' and (value < 0x00 or value > 0x48):
#             return "参数:“%s” 的值不在0~72之间，请检查后重试！" % key
#
#         # 判断：left_borders_size字段的值-》右边边框
#         if key == 'left_borders_size' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型的数据，请检查后重试！" % key
#         if key == 'left_borders_size' and value < 0x00:
#             return "参数:“%s” 的值不能为负整数，请检查后重试！" % key
#
#         # 判断：right_borders_size字段的值-》右边边框
#         if key == 'right_borders_size' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型的数据，请检查后重试！" % key
#         if key == 'right_borders_size' and value < 0x00:
#             return "参数:“%s” 的值不能为负整数，请检查后重试！" % key
#
#         # 判断：top_borders_size字段的值-》顶部边框
#         if key == 'top_borders_size' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型的数据，请检查后重试！" % key
#         if key == 'top_borders_size' and value < 0x00:
#             return "参数:“%s” 的值不能为负整数，请检查后重试！" % key
#
#         # 判断：bottom_borders_size字段的值-》底部边框
#         if key == 'bottom_borders_size' and not isinstance(value, int):
#             return "参数:“%s” 的值不是整数类型的数据，请检查后重试！" % key
#         if key == 'bottom_borders_size' and value < 0x00:
#             return "参数:“%s” 的值不能为负整数，请检查后重试！" % key
#
#     # 更新参数列表
#     copy_param_data = param_data.copy()
#     dict_data.update(copy_param_data)
#
#     #############################业务逻辑########################################
#
#     # 设置单元格背景颜色
#     self.__set_cells_colour(dict_data['fore_colour'])
#     # 设置字体样式
#     self.__set_font_style(dict_data['font_name'], dict_data['font_colour'],
#                           dict_data['font_bold'], dict_data['font_height'])
#     # 设置单元格边框
#     self.__set_cells_borders(dict_data['left_borders_size'], dict_data['right_borders_size'],
#                              dict_data['top_borders_size'], dict_data['bottom_borders_size'])
#
#     #############################业务逻辑########################################
#     return self.style


# set_cells_style = SetCellsStyle().set_cells_style_port(fore_colour =0,font_colour=1,font_name='name Times New Roman',font_bold=False,font_height=8,
#                                                        left_borders_size=9, right_borders_size=9, top_borders_size=9,bottom_borders_size=9)
# print(set_cells_style)