# -*- coding=utf-8 -*-

from common import LogUtil as logu


# properties处理类(*.properties)
class Properties:

    # 初始化并封装key：value
    def __init__(self, file_name):
        self.file_name = file_name
        self.properties = {}
        try:
            fopen = open(self.file_name, 'r', encoding='utf-8')
            for line in fopen:
                line = line.strip()
                if line.find('=') > 0 and not line.startswith('#'):
                    strs = line.split('=')
                    self.properties[strs[0].strip()] = strs[1].strip()
        except Exception as e:
            raise e
        else:
            fopen.close()

    # 是否存在key
    def has_key(self, key):
        return key in self.properties

    # 通过key获取值
    @logu.logExceptionHandler(log_func=logu.exceptionHandler)
    def get(self, key, default_value=''):
        if key in self.properties:
            return self.properties[key]
        return default_value

# 解析( 请使用ProjectTool.project_path + "/" + 文件名 )
@logu.logExceptionHandler(log_func=logu.exceptionHandler)
def parse(file_name):
    # 初始化
    return Properties(file_name)
