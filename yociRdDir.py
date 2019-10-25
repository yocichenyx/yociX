import os
import re

fileList = []

# Function can get *.xls/*.xlsx file from the directory
"""
dirpath: str, the path of the directory
"""


def _getfiles(dirPath):
    files = os.listdir(dirPath)
    # re match *.xls/xlsx
    ptn = re.compile('.*\.xlsx')
    for f in files:
        # isdir, call self
        if (os.path.isdir(dirPath + '\\' + f)):
            getfiles(dirPath + '\\' + f)
        # isfile, judge
        elif (os.path.isfile(dirPath + '\\' + f)):
            res = ptn.match(f)
            if (res != None):
                fileList.append(dirPath + '\\' + res.group())
        else:
            fileList.append(dirPath + '\\无效文件')


# Function called outside
"""
dirpath: str, the path of the directory
"""


def getfiles(dirPath):
    _getfiles(dirPath)
    return fileList


# if __name__ == "__main__":
#     path = 'D:\\pyfiles\\test'
#     res = getfiles(path)
#     print('提取结果：')
#     for f in res:
#         print(f)

# 10.17 实现了文件夹中特定类型文件的筛选（多层目录也可正常使用）
