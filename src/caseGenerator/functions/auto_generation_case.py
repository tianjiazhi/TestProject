# coding = utf-8
"""
@Time      : 2019/10/20 0020 19:45
@Author    : YunFan
@File      : auto_generation_case.py
@Software  : PyCharm
@Version   : 1.0
@Description: 自动生成测试用例
"""

from src.common.functions.read_write_excel import Reader
from src.common.functions.read_write_excel import SetCellsStyle
from src.common.functions.read_write_excel import NewFileWriter
from src.common.functions.log import logger
from get_file_path import get_join_path


log = logger

class WriteCellLogic():

    def __init__(self,w):
        self.write = w
        self.write_line = 0          # 初始化当前工作表的行数
        self.write_columns = 0       # 初始化当前工作表的列数

    def write_excel_title(self):
        """
        * 向工作簿下的每个工作表写入表头
        :return:
        """

        log.info("")
        log.info("向第{0}行写入标题开始》》》》》》》》》》》》》》》》》》》》》》》》》》》》》".format(self.write_line))

        self.write.write_data_cell(self.write_line, self.write_columns, "模块名称")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns , "模块名称"))

        self.write.write_data_cell(self.write_line, self.write_columns + 1, "分组名称")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 1, "分组名称"))

        self.write.write_data_cell(self.write_line, self.write_columns + 2, "用例名称")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 2, "用例名称"))

        self.write.write_data_cell(self.write_line, self.write_columns + 3, "关 键 字")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 3, "关 键 字"))

        self.write.write_data_cell(self.write_line, self.write_columns + 4, "请求地址")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 4, "请求地址"))

        self.write.write_data_cell(self.write_line, self.write_columns + 5, "请 求 头")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 5, "请 求 头"))

        self.write.write_data_cell(self.write_line, self.write_columns + 6, "请求参数")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 6, "请求参数"))

        self.write.write_data_cell(self.write_line, self.write_columns + 7, "检查字段")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 7, "检查字段"))

        self.write.write_data_cell(self.write_line, self.write_columns + 8, "期望结果")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 8, "期望结果"))

        self.write.write_data_cell(self.write_line, self.write_columns + 9, "执行状态")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 9, "执行状态"))

        self.write.write_data_cell(self.write_line, self.write_columns + 10, "实际结果")
        log.info("开始向【第{0}行第{1}列】写入标题【{2}】".format(self.write_line, self.write_columns + 10, "实际结果"))

        self.write_line = self.write_line + 1
        log.info("向第{0}行写入标题结束，进行换行后行号变成{1}《《《《《《《《《《《《《《《《《《".format(self.write_line -1,self.write_line))
        log.info("")




def run_write_excel_logic(line,row,sheet,write_cell_logic):
    """

    :param line: 读到当前行的内容
    :param row:  行数
    :param write_cell_logic:将读到的数据进行逻辑化处理后
    :return:
    """
    ##############################对模板中填写的信息进行非法法性校验####################################################

    if row != 0 and line[0].upper() not in ("YES","NO"):
        error = "数据源模板:工作表 {} 中的单元格 A{} 即[ifCheckValidity]字段的值【{}】系统未定义或未填写，请检查后重试！".format(sheet,row+1, line[0])
        log.error(error)
        raise ValueError(error)

    if row != 0 and line[1].upper() not in ("YES","NO"):
        error = "数据源模板:工作表 {} 中的单元格 B{} 即[ifCheckSpecial]字段的值【{}】系统未定义或未填写，请检查后重试！".format(sheet, row+1, line[1])
        log.error(error)
        raise ValueError(error)

    if row != 0 and line[2].upper() not in ("YES","NO"):
        error = "数据源模板:工作表 {} 中的单元格 C{} 即[ifCheckBoundary]字段的值【{}】系统未定义或未填写，请检查后重试！".format(sheet, row+1, line[2])
        log.error(error)
        raise ValueError(error)

    if row != 0 and not line[3]: # moduleName字段为必填项，当判定为空时，抛出错误。
        error = "数据源模板:工作表 {} 中的单元格 D{} 即[moduleName]字段的值【{}】不能为空，请填写后重试！".format(sheet, row+1, line[3])
        log.error(error)
        raise ValueError(error)

    if row != 0 and not line[4]: # interfaceName字段为必填项，当判定为空时，抛出错误。
        error = "数据源模板:工作表 {} 中的单元格 E{} 即[interfaceName]字段的值【{}】不能为空，请填写后重试！".format(sheet, row+1, line[4])
        log.error(error)
        raise ValueError(error)

    if row != 0 and not line[5]: # url字段为必填项，当判定为空时，抛出错误。
        error = "数据源模板:工作表 {} 中的单元格 F{} 即[url]字段的值【{}】不能为空，请填写后重试！".format(sheet, row+1, line[5])
        log.error(error)
        raise ValueError(error)

    if row != 0 and line[6].upper() not in ("POST","GET"): # method字段为必填项，当判定为空时，抛出错误。
        error = "数据源模板:工作表 {} 中的单元格 G{} 即[method]字段的值【{}】系统未定义或未填写，请填写后重试！".format(sheet, row+1, line[6])
        log.error(error)
        raise ValueError(error)

    ##############################根据模板中填写的信息进行合成测试用例####################################################

    if row == 0: # 当读到每个sheet页的行数为0时，执行写标题操作逻辑
        write_cell_logic.write_excel_title()
        return



    # if row != 0 and line[2].upper() == "YES" and line[8] and line[10] and (line[10] in line[8]) and line[11] and (line[13] or line[14]):
    #     """
    #         * 执行边界值测试时需要满足的条件：
    #         * ifCheckSpecial字段填写YES, params、checkParam、fieldType都不可为空，min字段和max字段的值至少有一个不为空，
    #         * 且在params字段中的值要包含在checkParam字段的值当中。
    #     """
    #     log.info("###########################开始执行边界值测试用例书写###############################")
    #
    #     if line[11] == "字符串":
    #
    #         log.info("*******************开始执行检查字段是“字符串”类型的边界值测试*************************")
    #
    #     if line[11] == "数字":
    #
    #         log.info("*******************开始执行检查字段是“数  字”类型的边界值测试*************************")
    #
    #     if line[11] == "手机号":
    #
    #         log.info("*******************开始执行检查字段是“手机号”类型的边界值测试*************************")
    #
    #
    # if  row != 0 and line[1].upper() == "YES" and line[8] and line[10] and (line[10] in line[8]) and line[11]:
    #     """
    #         * 执行特殊字符测试时需要满足的条件：
    #         * ifCheckBoundary字段填写 YES ,params、checkParam、fieldType都不可以为空。
    #         * 且在params字段中的值要包含在checkParam字段的值当中。
    #     """
    #     log.info("###########################开始执行特殊字符测试用例书写###############################")
    #
    #     if line[16]:
    #         log.info("***************需要测试的特殊字符，有值则允许输入该特殊字符***************")
    #     else:
    #         log.info("**************需要测试的特殊字符，无值则允许输入所有特殊字符**************")










# if row == 0:   # 首行
#     # 读取当前sheet页的总列数
#     cols = len(line)
#     log.info("工作表【{}】的首行共有{}列数据".format(sheet,cols))
#
#     # 将读取到的第一列列表数据写入到excel文件的第0行。
#     for index,value in enumerate(line):
#         log.info("开始向工作表【{0}】中写入表头【列数:{1}--->写入值:{2} 】".format(sheet,index,value))
#         # writer.write_data(row,index,value)
#         # 此时根据列数的索引做判断
#
# else:
#     # 功能描述：读指定单元格中的数据进行逻辑处理后并写入到其他行指定的单元格中
#
#     for index,value in enumerate(line):
#         log.info("开始向工作表【{0}】中写入测试用例数据【列数:{1}--->写入值:{2} 】".format(sheet,index,value))
#         writer.write_data(row,index,value)

def auto_generation_case(src_path,save_path):
    # 实例化一个Reader对象
    reader = Reader()
    # 打开待要读取的工作簿
    reader.open_excel(src_path)
    # 获取该工作簿下面所有的工作表
    sheetName = reader.get_sheets()

    # 实例化一个NewFileWriter对象
    writer = NewFileWriter()
    # 新建一个工作簿
    writer.new_workbook()

    for sheet in sheetName:
        reader.set_sheet(sheet)  # 通过sheet页面名称，读取所有的sheet页
        writer.new_sheet(sheet)  # 通过sheet页面名称，新建工作表

        # 在当前工作表页面实例化一个写入逻辑
        write_cell_logic = WriteCellLogic(writer)

        rows = reader.rows  # 读取当前sheet页的总行数
        log.info("工作表【{}】中共有【{}】行数据".format(sheet, rows))

        for row in range(rows):
            log.info("开始读取工作表【{0}】中的第【{1}】行数据...".format(sheet, row))
            line = reader.readline()  # 在当前sheet页分别读取每一行的数据
            log.info(line)

            # 根据读取到的行数据进行写逻辑
            run_write_excel_logic(line,row,sheet,write_cell_logic)

    writer.save_workbook(save_path)  # 对写入的文件进行保存


# 读取excel文件路径
src_path = get_join_path('files\\dataSrcFolder\\Interface_case_data_template.xls')
# 写入完成后excel文件保存的路径
save_path = get_join_path('files\\caseFolder\\Interface_case.xls')

auto_generation_case(src_path,save_path)

























