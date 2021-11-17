# -*- coding=utf-8 -*-
from common import ProjectCommon as tool
from common import Properties

properties = Properties.parse(tool.getPath() + "\\resources\\Settings.properties")

# 过滤器初始化大小
initialCapacity = int(properties.get("initialCapacity"))

# 过滤器错误率
errorRate = float(properties.get("errorRate"))

# excel表格路径
excelPlace = properties.get("excelPlace")

# sheet名称
sheetNames = properties.get("sheetNames")

# excel统一列索引
excelCellIndex = properties.get("excelCellIndex").split(',')

# 新建excel表格路径
excelCreatePath = properties.get("excelCreatePath")

# 新建excel表格Sheet名称(不重复的)
excelCreateSheetNameForNotExist = properties.get("excelCreateSheetNameForNotExist")

# 新建excel表格Sheet名称(重复的)
excelCreateSheetNameForExist = properties.get("excelCreateSheetNameForExist")

# excel表格基准定义
excelBenchmark = properties.get("excelBenchmark")

# excel文档的索引开始行
excelIndexStartRow = int(properties.get("excelIndexStartRow"))

# 判断索引1(用于判断存入为重复或不重复)
firstIndex = properties.get("firstIndex")

# 判断索引1(用于判断不重复的集合中，索引1为空并且筛选当前索引为不重复)
secondIndex = properties.get("secondIndex")