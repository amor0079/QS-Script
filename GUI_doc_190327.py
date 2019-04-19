#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/3/26 15:31
# @Author  : Amor
# @Site    : 
# @File    : GUI_doc.py
# @Software: PyCharm
from tkinter import *
from mailmerge import MailMerge
from datetime import date
import re

root = Tk()
root.title('ZBOE规格书自动生成脚本--Design by Amor')
root.iconbitmap('./icon.ico') #绝对路径或者其他
Label(root, text = "产品型号: ").grid(row = 0,column = 0)
Label(root, text = "规格书代码: ").grid(row = 1)
Label(root, text = "工程姓名(英文): ").grid(row = 2)
Label(root, text = "产品长(mm): ").grid(row = 3,column = 0)
Label(root, text = "宽: ").grid(row = 3,column = 2)
e1 = Entry(root,width = 15)
e1.grid(row = 0, column = 1, padx = 1, pady = 5)
e1.insert(0, "ZDP- ")
e2 = Entry(root,width = 15)
e2.grid(row = 1, column = 1, padx = 10, pady = 5)
e2.insert(0, "ZBOE ")
e3 = Entry(root,width = 15)
e3.grid(row = 2, column = 1, padx = 10, pady = 5)
e3.insert(0, '')
e4 = Entry(root,bd = 2,width = 15) #bd 线框深度
e4.grid(row = 3, column = 1,sticky =  N, padx = 5, pady = 5)
e5 = Entry(root,bd = 2,width = 15)
e5.grid(row = 3, column = 3, sticky =N,padx = 5, pady = 5)

def show():
    template = "./template_ZBOE_data_sheet.docx"
    document = MailMerge(template)
    a1= e1.get()
    colour_len =a1.split('-')[1]
    if len(re.sub('[^a-zA-Z]','',colour_len)) ==1:
        sm_colour = "single-color"
    elif len(re.sub('[^a-zA-Z]','',colour_len)) >=2:
        sm_colour = "multi-color"
    up_sm_colour = sm_colour.upper()

    if int(re.sub('[a-zA-Z]','',colour_len)[-1])/ 2 ==0:
        jixing = 'Multiplex  Common  Anode'
    else:
        jixing = 'Multiplex  Common  Cathode'

    # colour = {'G':'Yellow Green','M':'Super Bright Yellow Green','Y':'Yellow','T':'Super Bright Yellow','R':'Red',
    #           'S':'Super Bright Red','D':'Super Bright Red','W':'Super Bright Red','E':'Oright',
    #           'Q':'Super Bright Orange',
    #           'B':'Super Bright Blue','A':'Super Bright Amber','V':'Super Bright Green','X':'Super Bright Green',
    #           'C':'Super Bright White'
    #           }
    #
    # k =re.sub('[a-zA-Z]', '', colour_len)
    # for key in colour.keys():
    #     if k[0] == key:
    #         co1= colour[key]
    #     elif k[1] ==key:
    #         co2= colour[key]

    #
                # elif len(re.sub('[^a-zA-Z]', '', colour_len)) == 3:
                #     for key in colour:
                #         if re.sub('[a-zA-Z]','',colour_len)[0] == key:
                #             colour_1 = colour[key]
                #         if re.sub('[a-zA-Z]','',colour_len)[1] ==key:
                #             colour_2 = colour[key]
                #         if re.sub('[a-zA-Z]','',colour_len)[2] ==key:
                #             colour_3 = colour[key]

    cust_1 = {'name':e1.get(),
              'date1':'{:%B %d,%Y}'.format(date.today()),
              'date2':'{:%Y-%m-%d}'.format(date.today()),
              'spec_no':e2.get(),
              'engineer':e3.get(),
              'wide':e4.get(),
              'length':e5.get(),
              'n_colour':sm_colour,
              'up_sm_colour':up_sm_colour,
              'jixing':jixing,
              # 'colour_1':co1,
              # 'Colour_2':co2

              }
    document.merge_pages([cust_1])
    # print(document.get_merge_fields())
    document.write("./{}.docx".format(e1.get()))
    e1.delete(0, END)

Button(root, text = "确认输出文档", width = 10, command = show)\
             .grid(row = 4, column = 0, sticky = W, padx = 10, pady =5)
Button(root, text = "退出", width = 10, command = root.quit)\
             .grid(row = 4, column =3, sticky = E, padx = 10, pady = 5)

mainloop()
