# -*- coding:UTF-8 -*-

import os,sys

PROJECT_NAME = "ExcelFieldFilter"
project_path = os.path.abspath(os.path.dirname(__file__))
root_path = project_path[:project_path.index(PROJECT_NAME) + len(PROJECT_NAME)]

# root_path = os.path.dirname(os.path.realpath(sys.executable))

def getPath():
    return root_path

#判断文档是否存在
def FileExist(fPath):
    if not os.path.exists(fPath):
        return False
    return True

def sortFileNameList(fileNameList, benchmark):
    result = []
    # 先把mark放入result
    for fileName in fileNameList:
        # splitPart = fileName.split('-')
        if benchmark in fileName:
            result.append(fileName)
            fileNameList.remove(fileName)

    # 对剩余的名称排序后放入result中并返回
    fileNameList.sort(key = lambda i:len(i), reverse=True)
    for file in fileNameList:
        result.append(file)
    return result