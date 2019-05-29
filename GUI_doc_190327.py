#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/3/26 15:31
# @Author  : Amor
# @Site    : 
# @File    : GUI_doc.py
# @Software: PyCharm
from tkinter import *
from tkinter import messagebox
from mailmerge import MailMerge
from datetime import date
import re

#晶片字典
chip_Inf = {
    'G': ('G', 'Yellow Green', 'GaP/GaP', 2.2, 2.7, 570, 30),
    'M': ('M', 'Super Bright Yellow Grenn ', 'AlGalnP/GaAs', 2.0, 2.5, 572, 15),
    'Y': ('Y', 'Yellow', 'GaAsP/GaP', 2.0, 2.7, 588, 30),
    'T': ('T', 'Super Bright Yellow', 'AlGalnP/GaAs', 2.0, 2.5, 588, 15),
    'R': ('R', 'Red', 'GaP/GaP', 2.0, 2.7, 640, 90),
    'S': ('S', 'Super Bright Red', 'AlGaAs/GaAs', 2.0, 2.5, 640, 20),
    'D': ('S', 'Super Bright Red', 'AlGaAs/GaAs', 2.0, 2.5, 640, 20),
    'W': ('W', 'Super Bright', 'AlGalnP/GaAs', 2.0, 2.5, 630, 20),
    'E': ('E', 'Orange', 'GaAsP/GaP', 2.2, 2.7, 620, 35),
    'Q': ('Q', 'Super Bright Orange', 'AlGalnP/GaAs', 2.0, 2.5, 625, 20),
    'B': ('B', 'Super Bright Blue', 'InGaN/GaN', 3.5, 4.0, 468, 25),
    'A': ('A', 'Amber', 'AlGalnP/GaAs', 2.0, 2.5, 607, 20),
    'V': ('V', 'Super Bright Green', 'InGaN/GaN', 3.5, 4.0, 525, 35),
    'X': ('V', 'Super Bright Green', 'InGaN/GaN', 3.5, 4.0, 535, 35),
    'C': ('C', 'Super Bright White', 'InGaN/GaN', 3.2, 3.8)}


#筛选晶片函数
def name_resolution():
    global chip, chip_len, polarity, tem_choice
    a1 = e1.get() #获取输入
    tem_choice = a1.split('-')[0][-1] #依据最后一位来选择模板

    colour_len_split = a1.split('-')[1] #使用split函数分隔输入信息
    chip = re.sub('[^a-zA-Z]', '', colour_len_split)
    # select_new_chip = list(chip)

    chip_len = len(chip) #长度 判断晶片颜色个数

    polarity = int(re.sub('[a-zA-Z]', '', colour_len_split)[-1]) #极性


#筛选晶片函数
def col_choice():
    global choice_chip_inf
    choice_chip_inf = []
    for i in chip: #遍历颜色
        for key, value in chip_Inf.items(): #遍历chip字典
            if i == key:
               choice_chip_inf.append(value)

#晶片单复数函数
def sin_mul():
    global sm_colour
    if chip_len ==1:
        sm_colour = "single-color"
    elif chip_len >=2:
        sm_colour = "multi-color"

def jx():
    global jixing
    if polarity % 2 ==0:  #判断极性
        jixing = 'Multiplex  Common  Anode'
    else:
        jixing = 'Multiplex  Common  Cathode'

def message():
    messagebox.showinfo(title= 'Congratulation', message='规格书创建成功')


#选择模板函数
def choice_tem():
    #模板清单
    tem0_ZDS = r".\ds_tem\tem0_zds.docx"
    tem1_ZDD = r".\ds_tem\tem1_zdd.docx"
    tem2_ZDT = r".\ds_tem\tem2_zdt.docx"
    tem3_ZDF = r".\ds_tem\tem3_zdf.docx"
    tem4_ZDP_1 = r".\ds_tem\tem4_zdp_1.docx"
    tem5_ZDP_2 = r".\ds_tem\tem5_zdp_2.docx"
    tem6_ZDP_3 = r".\ds_tem\tem6_zdp_3.docx"
    tem7_ZDP_4 = r".\ds_tem\tem7_zdp_4.docx"
    #判断
    sin_mul() #判断晶片单复数。
    jx() #判断晶片极性
    up_sm_colour = sm_colour.upper() #颜色大写

    if tem_choice == 'S': #选择ZDS模板
        document = MailMerge(tem0_ZDS) #选择模板
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'D': #选择ZDD模板
        document = MailMerge(tem1_ZDD)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == "T": #选择ZDT模板
        document = MailMerge(tem2_ZDT)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'F':
        document = MailMerge(tem3_ZDF)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'P' and chip_len == 1:
        document = MailMerge(tem4_ZDP_1)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'P' and chip_len == 2:
        document = MailMerge(tem5_ZDP_2)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'P' and chip_len == 3:
        document = MailMerge(tem6_ZDP_3)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'P' and chip_len == 4:
        document = MailMerge(tem7_ZDP_4)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': e2.get(),
                  'engineer': e3.get(),
                  'wide': e4.get(),
                  'length': e5.get(),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  # 'colour_1':co1,
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))
    global template
    #需要加入try except

def colour(): #判断颜色个数赋值晶片参数
    if chip_len == 1:
        colour1 = choice_chip_inf[0][0]
        pass    #晶片其他参数
    elif chip_len == 2:
        colour1 = choice_chip_inf[0][0]
        colour2 = choice_chip_inf[1][0]
    elif chip_len == 3:
        colour1 = choice_chip_inf[0][0]
        colour2 = choice_chip_inf[1][0]
        colour3 = choice_chip_inf[2][0]
    elif chip_len == 4:
        colour1 = choice_chip_inf[0][0]
        colour2 = choice_chip_inf[1][0]
        colour3 = choice_chip_inf[2][0]
        colour4 = choice_chip_inf[3][0]
    #需要加入try except


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
    name_resolution() #执行名字分解读取
    col_choice() #遍历晶片字典,赋值对应的
    colour() #赋值晶片参数
    choice_tem()  # 执行选择模板函数
    message()  #弹框确认
  # e1.delete(0, END) #清空el中输入框中的内容

Button(root, text = "确认输出文档", width = 10, command = show)\
             .grid(row = 4, column = 0, sticky = W, padx = 10, pady =5)
Button(root, text = "退出", width = 10, command = root.quit)\
             .grid(row = 4, column =3, sticky = E, padx = 10, pady = 5)

mainloop()