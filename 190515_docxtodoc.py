#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/5/15 9:17
# @Author  : Amor
# @Site    : 
# @File    : 190515_docx to doc.py
# @Software: PyCharm

# import os
# from win32com import client as wc
#
# path = 'D:\\30Day\\30天自动办公训练营\\script\\qiusheng-Script\\ZDP-23232C-40E生产条件票\\'
# files = []
# for file in os.listdir(path):
#     if file.endswith('.docx'):
#         files.append(path + file)
#
# word = wc.Dispatch('Word.Application')
# for file in files:
#     doc = word.Documents.Open(file) #打开word文件
#     doc.SaveAs(file[0:-1], 16)#另存为后缀为".docx"的文件，其中参数12指docx文件 参数16指doc文件
#     doc.Close() #关闭原来word文件
# word.Quit()
# print("完成！")
import os
from win32com import client
from tkinter import *  # 导入 Tkinter 库
from tkinter.filedialog import askdirectory


# 第1步，实例化object，建立窗口window

def doc_to_docx(path):
    if os.path.splitext(path)[1] == ".docx":
        word = client.Dispatch('Word.Application')
        doc = word.Documents.Open(path)  # 目标路径下的文件
        doc.SaveAs(os.path.splitext(path)[0] + ".doc", 16)  # 转化后路径下的文件
        doc.Close()
        word.Quit()


def find_file(path, ext, file_list=[]):
    dir = os.listdir(path)
    for i in dir:
        i = os.path.join(path, i)
        if os.path.isdir(i):
            find_file(i, ext, file_list)
        else:
            if ext == os.path.splitext(i)[1]:
                file_list.append(i)
    return file_list


def selectPath():
    path_ = askdirectory()
    cpath.set(path_)
    path = path_
    ext = '.docx'
    file_list = find_file(path, ext, file_list=[])
    for file in file_list:
        doc_to_docx(file)


root = Tk()
root.title("Docx To Doc")
root.iconbitmap('./icon.ico')
root.geometry('350x50')
cpath = StringVar()
Label(root, text="目标路径:").grid(row=0, column=0)
Entry(root, textvariable=cpath).grid(row=0, column=1)
Button(root, text="选择文件夹", command=selectPath).grid(row=0, column=3)
root.mainloop()