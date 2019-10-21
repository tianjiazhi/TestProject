# -*- encoding: utf-8 -*-
"""
@File        : log.py
@Time        : 2019/10/18 0018 下午 14:24
@Author      : YunFan
@Email       : tianjiazhi615@163.com
@Software    : PyCharm
@Version     : 1.0
@Description : 
"""
import os
import logging
from logging.handlers import TimedRotatingFileHandler
from get_file_path import get_join_path


# 存放log文件的绝对路径
logPath = get_join_path("files\\logFolder")

class Logger(object):

    def __init__(self, logger_name='logs…'):

        self.logger = logging.getLogger(logger_name)
        logging.root.setLevel(logging.NOTSET)
        # 日志文件的名称
        self.log_file_name = 'logs'
        # 最多存放日志的数量
        self.backup_count = 10
        # 日志输出级别
        # self.console_output_level = 'WARNING'      # 打印在控制台日志级别
        self.console_output_level = 'DEBUG'          # 打印在控制台日志级别
        self.file_output_level = 'DEBUG'             # 写在日志文件级别
        # self.file_output_level = 'WARNING'         # 写在日志文件级别
        # 日志输出格式
        self.formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    def get_logger(self):
        """在logger中添加日志句柄并返回，如果logger已有句柄，则直接返回"""

        if not self.logger.handlers:  # 避免重复日志

            console_handler = logging.StreamHandler()
            console_handler.setFormatter(self.formatter)
            console_handler.setLevel(self.console_output_level)
            self.logger.addHandler(console_handler)


            # 每隔一个小时重新创建一个日志文件，最多保留backup_count份
            file_handler = TimedRotatingFileHandler(filename=os.path.join(logPath,self.log_file_name), when='H',
                                                    interval=1, backupCount=self.backup_count, delay=True,
                                                    encoding='utf-8')
            file_handler.setFormatter(self.formatter)
            file_handler.setLevel(self.file_output_level)
            self.logger.addHandler(file_handler)

        return self.logger

logger = Logger().get_logger()


