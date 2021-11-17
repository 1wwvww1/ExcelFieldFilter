# -*- coding=utf-8 -*-

import logging.config, logging as log
import os, yaml, functools
from common import ProjectCommon as tool

root_path = tool.getPath()

def setup_logging(default_path=(root_path+'\\resources\\logging.yaml'), default_level=logging.DEBUG):

    path = default_path
    if os.path.exists(path):
        with open(path, 'rt') as f:
            config = yaml.load(f.read(),Loader=yaml.FullLoader)
        logging.config.dictConfig(config)
    else:
        logging.basicConfig(level=default_level)
        print('logging.yaml doesn\'t exist,log open fail')


def logExceptionHandler(log_func, isContinue=True):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kw):

            result = None
            try:

                result = func(*args, **kw)
            except Exception as ex:
                # 处理异常
                if(isContinue):
                    log_func("call function: %s(), args: %s, kw: %s, raise Excepation " % (
                        func.__name__, args, kw), ex)
                else:
                    raise ex

                # input("打印过程中出现错误，回车继续任务:")
            return result
        return wrapper
    return decorator


def exceptionHandler(exText, ex):
    message = exText+str(ex)+"\n"
    log.error(message, exc_info = True)
    print(message)