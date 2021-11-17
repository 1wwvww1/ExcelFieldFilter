
import os, logging as log,random
from service import ExcelService as excelInstaller
from entity import ConfigEntity as config, ErroMessageEntity as erro
from common import ProjectCommon as comm, BloomFilter as bloom, LogUtil as logu

# 日志开启
logu.setup_logging()

@logu.logExceptionHandler(log_func=logu.exceptionHandler, isContinue=False)
def useBloomFilter(excelInstance, cnBloomFilter, newExcelDataMap, newExcelDataMapE, isBenchmark=False):

    # 获取对应的sheet并遍历索引栏获取对应要获取的索引在第几列
    sheet = excelInstance.loadSheetName(config.sheetNames)
    indexMap = excelInstance.readExcelIndexToMap(config.excelCellIndex, config.excelIndexStartRow, sheet)

    indexCell = indexMap[config.firstIndex]
    bloomFilter = cnBloomFilter
    # 获取对于索引列的所有内容字典(key = 第几行， value = 索引列对应的内容）
    indexColMap = excelInstance.readExcelCellToMap(indexCell, sheet)

    # 遍历所有字串是否存在于布隆过滤器中
    sheetCellValue = list(indexColMap.values())
    sheetCellKey = list(indexColMap.keys())
    for sheetCellIndex in range(0, sheetCellValue.__len__()):
        cellValue = str(sheetCellValue[sheetCellIndex]).strip()
        # 支持如果当前为空那么就添加到不重复的集合中
        if("".__eq__(cellValue) or "None".__eq__(cellValue) or cellValue == {None}):
            # 重新构建key
            cellValue = "YTSNVGH"+str(sheetCellIndex)+str(random.randrange(1,config.initialCapacity))
            cellIndex = indexMap[config.firstIndex]
            # field = sheet.cell(row=sheetCellKey[sheetCellIndex], column=cellIndex).value
            newExcelData = excelInstance.readExcelOneRow(indexMap, sheetCellKey[sheetCellIndex], sheet)
            newExcelDataMap[cellValue] = newExcelData
            continue

        # 如果是基准表就直接插入不需要查询
        if isBenchmark:
            bloom.insert_bloomFilter(cellValue, bloomFilter, False)
        else:
            isNotBloom = bloom.insert_bloomFilter(cellValue, bloomFilter, True)
            # 非基准表添加失败表示存在
            if not isNotBloom:
                # 如果存在过滤器中就去查看'不存在'的缓存是否有数据，如果有清除，并放入'存在'的缓存中
                if newExcelDataMap.__contains__(cellValue):
                    newExcelData = newExcelDataMap.pop(cellValue)
                    newExcelDataMapE[cellValue] = newExcelData

                # 如果不在'不存在'的缓存中，就去查询‘存在’的缓存是否有内容。如果没有就添加到‘不存在’的缓存，否则跳过
                else:
                    if not newExcelDataMapE.__contains__(cellValue):
                        cellIndex = indexMap[config.cnIndex]
                        newExcelData = excelInstance.readExcelOneRow(indexMap, sheetCellKey[sheetCellIndex], sheet)
                        newExcelDataMap[cellValue] = newExcelData
            else:
                # 如果不存在于布隆过滤器中就直接获取数据存入'不存在'的缓存中
                cellIndex = indexMap[config.firstIndex]
                # field = sheet.cell(row=sheetCellKey[sheetCellIndex], column=cellIndex).value
                newExcelData = excelInstance.readExcelOneRow(indexMap, sheetCellKey[sheetCellIndex], sheet)
                newExcelDataMap[cellValue] = newExcelData

if __name__ == '__main__':

    # 先检查文件个数是否小于2
    excelPath = config.excelPlace
    if comm.FileExist(excelPath) :
        os.chdir(excelPath)
        if os.listdir(excelPath).__len__() < 2:
            raise Exception(erro.excelFileIsLess)
    else:
        raise Exception(erro.excelPathNotExist)

    # 获取路径下面第一层目录的所有文件
    excelFilePathList = []
    for dirpath, dirnames, filenames in os.walk(excelPath):
        excelFilePathList = [os.path.join(dirpath, names) for names in filenames]

    for path in excelFilePathList:
        if not comm.FileExist(path):
            raise Exception(erro.excelFileIsNotExist)

    # 进行文件排序(基准表 > 文件名称)
    excelFilePathList = comm.sortFileNameList(excelFilePathList, config.excelBenchmark)

    # 循环创建对应的excel实例
    excelInstances = []
    for excelFilePath in excelFilePathList:
        try:
            excel = excelInstaller.ExcelService(excelFilePath)
            excelInstances.append(excel)
        except Exception as ex:
            raise ex

    log.info("共创建了 {count} 个Excel实例！".format(count = excelInstances.__len__()))

    # 布隆过滤器
    cnBloomFilter = None
    try:
        cnBloomFilter = bloom.construct_bloomFilter(config.initialCapacity, config.errorRate)
    except Exception as ex:
        raise ex

    # 先在过滤器中添加基准表的数据
    newExcelDataNotExist = {}
    newExcelDataExist = {}
    # useBloomFilter(excelInstances[0], cnBloomFilter, newExcelDataNotExist, newExcelDataExist, True)
    # excelInstances.pop(0)

    for excelInstance in excelInstances:
        useBloomFilter(excelInstance, cnBloomFilter, newExcelDataNotExist, newExcelDataExist, False)

    # print(newExcelDataExist)

    # 过滤之后进行创建新表格
    newExcelDataList = []
    newExcelDataList.append( list(newExcelDataNotExist.values()) )
    newExcelDataList.append( list(newExcelDataExist.values()) )

    newExcelSheetList = []
    newExcelSheetList.append(config.excelCreateSheetNameForNotExist)
    newExcelSheetList.append(config.excelCreateSheetNameForExist)

    excelInstaller.createExcelBook(config.excelCreatePath, newExcelSheetList,
                                   config.excelCellIndex, newExcelDataList)

    # TODO 多线程执行并添加行数据