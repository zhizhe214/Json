# -*- coding: utf-8 -*-
import xlrd
import json
import collections
import os

# xlwt 写入 Excel
# xlrd 读取 Excel
textName='大数据编码'
absolutePath=''      # 打开路径
excelName=''         # Excel 文件名称
totalRows=0          # 总的有效行数
totalColumns=0       # 总的有效列数
totalSystems=0       # 筛选出来的系统总数
systemsList=[]
belongRowAndColumn=()       # “Belong” 所在的行和列的索引

 # 四个 标题 的列的索引
belongColumn=0                  # “Belong” 所在的列的索引
tagTitleColumn=0                # “TagTitle” 所在的列的索引
unitColumn=0                    # “Unit” 所在的列的索引
tagGroupColumn=0                # “TagGroup” 所在的列的索引

totalRows=0     # 总行数
totalColumns=0  # 总列数
parentDir=''

def getParentDirectory(file):
    global parentDir
    parentDir=os.path.dirname(absolutePath)
    print('父级目录： '+parentDir)
    return parentDir


 # 判断字符串: 不为空
def stringIsNull(str):
    if len(str)>0:
        return True

# 创建 Text文件
def CreateText(fileName):
    getParentDirectory(absolutePath)
    print('创建文件： '+parentDir+'\\'+fileName.replace('.xls','')+".txt")
    global textPath
    textPath=parentDir+'\\'+fileName.replace('.xls','')+".txt"
    print('textPath： '+textPath)
    file=open(textPath,'w+')    # w+表示：打开一个文件用于读写，如果文件已经存在则将其覆盖，如果文件不存在则创建文件。
    file.close()
# 追加
def AddToTextContent(message):
    file=open(textPath,'a+')
    file.write(message)
    file.close()


def openExcelAndSetValue():
    print('打开Excel： '+absolutePath)
    book = xlrd.open_workbook(absolutePath)
    global sheet
    sheet = book.sheet_by_index(0)  # 根据 sheet 索引操作
    # sheet=book.sheet_by_name('sheet1')  # 根据 sheet 名称操作
    global totalRows
    rowCount=sheet.nrows        # excel里面有多少行
    totalRows=rowCount
    global totalColumns
    columnCount=sheet.ncols     # excel里面有多少列
    totalColumns=columnCount
    print('有效总行数：%d'%totalRows)
    print('有效总列数：%d'%totalColumns)


# print('该Excel一共 %d 行' %rowCount)
# print('该Excel一共 %d 列'%columnCount)

# 获取  行、列、单元格  的内容
# print(sheet.cell(0,0).value)    #获取第 1 行第 1 列（即：某个单元格）的值
#
# print(sheet.row_values(0))      # 获取 一整行 的内容
# print(sheet.col_values(2))      # 获取 一整列 的内容

'''
# 循环获取每行的内容
for i in range(sheet.nrows):
    # print(sheet.row_values(i))
    AddToTextContent(str(sheet.row_values(i)))   # 将 Excel 每行内容，遍历并追加包 Text 中
'''

# 查找内容位于哪一行和哪一列
def FindCellByName(name):
    thisRow=0
    thisColumn=0
    for r in range(totalRows):
        for c in range(totalColumns):
            cellValue=sheet.cell(r,c).value
            if  str(cellValue)==name:
                thisRow=r
                thisColumn=c
    return thisRow,thisColumn   # 返回行和列的索引

# 求得 “ 'Belong'、'TagTitle'、'Unit'、'TagGroup' ” 的列编号
def getColumnId():
    global belongColumn
    belongColumn=FindCellByName('Belong')[1]
    print('belongColumn： %d'%belongColumn)
    global tagTitleColumn
    tagTitleColumn=FindCellByName('TagTitle')[1]
    global unitColumn
    unitColumn=FindCellByName('Unit')[1]
    global tagGroupColumn
    tagGroupColumn=FindCellByName('TagGroup')[1]

# 求得几大系统的名称列表
def getSystemsList():
    for item in range(1, totalRows):
        cellValue = sheet.cell(item, belongColumn).value
        if stringIsNull(cellValue):
            if cellValue not in systemsList:  # 判断 元素 是否在列表list中
                systemsList.append(cellValue)
                # print(cellValue)
    print("系统 List 包含：" + str(systemsList))







# 声明字典变量的两种方式
# belongSystemDict={}    # 声明字典变量（默认的创建方法）
belongSystemDict=collections.OrderedDict()  # 声明字典变量（会按照输入的顺序排序）

totalList=[]


 # 获取某一个系统的 value
def getOneBelongSystemValue(titleName):
    mydata=collections.OrderedDict()
    myList=[]
    for r in range(1,totalRows):
        if  sheet.cell(r,belongColumn).value==titleName:
            valueTagTitle=sheet.cell(r,tagTitleColumn).value
            valueUnit=sheet.cell(r,unitColumn).value
            valueTagGroup=int(sheet.cell(r,tagGroupColumn).value)

            info =collections.OrderedDict()
            info["Id"] = str(valueTagGroup)
            info["Name"] = valueTagTitle
            info["Unit"] = valueUnit
            myList.append(info)
            mydata["BelongSystem"] = titleName
            mydata["List"] = myList
    return mydata




def exportTextWithJson():
    for i in range(len(systemsList)):
        global totalList
        totalList.append(getOneBelongSystemValue(systemsList[i]))
    jsonStr = json.dumps(totalList, ensure_ascii=False)
    print(jsonStr)
    CreateText(excelName)
    AddToTextContent(jsonStr)