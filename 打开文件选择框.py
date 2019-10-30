import win32ui
# -*- coding: utf-8 -*-
def openFileDialog():
    dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
    dlg.SetOFNInitialDir(r'G:\迅雷下载')  # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    pathname = dlg.GetPathName()  # 获取选择的文件名称
    filename = dlg.GetFileName() # 获取选择的文件名称
    # print('文件路径为：%s'%pathname)
    # print('文件名称为：%s'%filename)
    return pathname,filename
# self.lineEdit_InputId_AI.setText(filename)  #将获取的文件名称写入名为“lineEdit_InputId_AI”可编辑文本框中
