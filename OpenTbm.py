# -*- coding: utf-8 -*-
import tkinter
import 打开文件选择框
import 用Python完成Excel的常用操作

root=tkinter.Tk()
root.geometry('600x400+500+350')
root.title('盾构大数据点位生成')

def getfilePathAndName():
    fileAndPathList=打开文件选择框.openFileDialog()
    用Python完成Excel的常用操作.absolutePath= fileAndPathList[0]
    用Python完成Excel的常用操作.excelName= fileAndPathList[1]
    print('路径为：%s' % 用Python完成Excel的常用操作.absolutePath)
    print('名称为：%s' % 用Python完成Excel的常用操作.excelName)
    用Python完成Excel的常用操作.openExcelAndSetValue()
    用Python完成Excel的常用操作.getColumnId()

    用Python完成Excel的常用操作.getSystemsList()
    用Python完成Excel的常用操作.exportTextWithJson()

tuyaBtn=tkinter.Button(root,text='土压平衡点位配置',bg='#d9d6c3',activebackground='yellow',highlightbackground='red',
                       width=20,height=2,font=('微软雅黑','15','bold '),
                       command=getfilePathAndName
                       )
tuyaBtn.pack()
def on_enter(e):
    tuyaBtn['background'] = 'green'
def on_leave(e):
    tuyaBtn['background'] = '#d9d6c3'
# 绑定事件；
tuyaBtn.bind('<Enter>',on_enter)
tuyaBtn.bind('<Leave>',on_leave)
nishuiBtn=tkinter.Button(root,text='泥水平衡点位配置',bg='#d9d6c3',width=20,height=2,font=('微软雅黑','15','bold'))
nishuiBtn.pack()
OpenTbmBtn=tkinter.Button(root,text='敞开式TBM点位配置',bg='#d9d6c3',width=20,height=2,font=('微软雅黑','15','bold'))
OpenTbmBtn.pack()
root.mainloop()