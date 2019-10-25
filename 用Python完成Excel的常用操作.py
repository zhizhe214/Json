import xlrd
import json
import collections

# xlwt 写入 Excel
# xlrd 读取 Excel
textName='大数据编码'
absolutePath='E:\\python练习\\操作excel\\json序列\\'+ textName + '.txt'
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


 # 判断字符串: 不为空
def stringIsNull(str):
    if len(str)>0:
        return True

# 创建 Text文件
def CreateText(fileName):
    global textName
    textName=fileName
    file=open(absolutePath,'w')    # w+表示：打开一个文件用于读写，如果文件已经存在则将其覆盖，如果文件不存在则创建文件。
    print('创建文件： '+fileName+".txt")
    file.close()
# 追加
def AddToTextContent(message):
    file=open(absolutePath,'a+')
    file.write(message)
    file.close()


book=xlrd.open_workbook('点位模板dl520.xls')
sheet=book.sheet_by_index(0)    # 根据 sheet 索引操作
# sheet=book.sheet_by_name('sheet1')  # 根据 sheet 名称操作

rowCount=sheet.nrows        # excel里面有多少行
totalRows=rowCount
columnCount=sheet.ncols     # excel里面有多少列
totalColumns=columnCount
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




belongColumn=FindCellByName('Belong')[1]
tagTitleColumn=FindCellByName('TagTitle')[1]
unitColumn=FindCellByName('Unit')[1]
tagGroupColumn=FindCellByName('TagGroup')[1]


for item in range(1,totalRows):
    cellValue=sheet.cell(item,belongColumn).value
    if stringIsNull(cellValue) :
        if cellValue not in systemsList:    # 判断 元素 是否在列表list中
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
            # belongSystem = {}
            # belongSystem["BelongSystem"] = data


            # print(str(valueTagGroup)+' , '+valueTagTitle+' , '+valueUnit)
'''
# for i in range(len(systemsList)):
#     # getOneBelongSystemValue('刀盘系统')
#     getOneBelongSystemValue(systemsList[i])
'''

for i in range(0,5):
    totalList.append(getOneBelongSystemValue(systemsList[i]))
    # totalList.append(data)

jsonStr = json.dumps(totalList, ensure_ascii=False)
print(jsonStr)
CreateText('生成大数据JSON序列')
AddToTextContent(jsonStr)