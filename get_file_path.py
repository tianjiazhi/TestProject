# coding = utf-8
"""
@Time      : 2019/10/19 0019 23:53
@Author    : YunFan
@File      : get_file_path.py
@Software  : PyCharm
@Version   : 1.0
@Description: python解释器的版本号：3.7.4
"""
import os

def get_project_path():
    # 获取当前文件存放的绝对路径
    abspath = os.path.abspath(__file__)
    # 获取当前文件的路径名称即项目路径
    dirname = os.path.dirname(abspath)
    return dirname


def get_join_path(relative_path):
    project_path = get_project_path()
    # 将项目路径和相对路径拼接成为完整的绝对路径并返回
    join_path = os.path.join(project_path,relative_path)
    return join_path

