# Date: 2022/05/05
# Author: Bernstein

import re
import os
import xlwt
import xlrd
# 将腾讯会议昵称分离
# 改进建议：可以使用正则匹配进行改进
path = os.getcwd()
model = input('请选择模式：1.批处理并保存/2.单独处理并保存/3.单独处理并输出结果 (1 or 2 or 3)(暂未开通)')
model = input('请选择模式：1.txt文件/2.xlsx文件 (1 or 2)(暂未开通)')
def OpenFile():
    inputFileName = input('请输入当前文件夹下需要分离昵称的完整文件名：')
    filepath = os.path.join(path, inputFileName)
    return filepath
def CreateFile():
    textbook = xlwt.Workbook(encoding='utf-8')
    return textbook
def outputProcess():
    while True:
        filepath = OpenFile()
        outputFileName = input('请输入一个目标表格文件名，此文件将保存在当前文件夹下：')
        textbook = CreateFile()
        sheet = textbook.add_sheet('名单')
        
        with open(filepath,'r',encoding='utf-8') as file:
            lenOfLine = 0
            for line in file:
                for i in range(len(line)):
                    if line[i] == '(':
                        sheet.write(lenOfLine, 0, line[i+1:-2])
                        print(line[i+1:-2])
                        break
                lenOfLine += 1
        textbook.save(path + '\\' + outputFileName +'.xlsx')
        
outputProcess()