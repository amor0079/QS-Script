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
    'W': ('W', 'Super Bright Red', 'AlGalnP/GaAs', 2.0, 2.5, 630, 20),
    'E': ('E', 'Orange', 'GaAsP/GaP', 2.2, 2.7, 620, 35),
    'Q': ('Q', 'Super Bright Orange', 'AlGalnP/GaAs', 2.0, 2.5, 625, 20),
    'B': ('B', 'Super Bright Blue', 'InGaN/GaN', 3.5, 4.0, 468, 25),
    'A': ('A', 'Amber', 'AlGalnP/GaAs', 2.0, 2.5, 607, 20),
    'V': ('V', 'Super Bright Green', 'InGaN/GaN', 3.5, 4.0, 525, 35),
    'X': ('V', 'Super Bright Green', 'InGaN/GaN', 3.5, 4.0, 535, 35),
    'C': ('C', 'Super Bright White', 'InGaN/GaN', 3.2, 3.8, '/' , '/')}


#筛选晶片函数
def name_resolution():
    global chip, chip_len, polarity, tem_choice, ink_choice, inch_choice
    a1 = e1.get() #获取输入
    tem_choice = a1.split('-')[0][-1] #依据最后一位来选择模板

    ink_choice = a1.split('-')[2][0] #依据第三段第一位来判断油墨颜色。

    inch_choice = a1.split('-')[1][1] #提取第二段第二位

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

def ink(): #判断油墨颜色
    global ref_ink
    if ink_choice == '2':
        ref_ink = 'black'
    elif ink_choice == '1':
        ref_ink = 'gray'
    return ref_ink

def power():  #增加功率参数
    global pw
    pw = int(float(co1_VF_m) * 20)
    return pw

def inch():
    global inc, millimeter
    inc = '0.{}'.format(inch_choice)
    millimeter = str(int(inch_choice)/10 * 25.4)
    return inc, millimeter



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
    ink() #判断塑壳油墨
    power() #判断功率
    inch() #赋值尺寸 英寸
    up_sm_colour = sm_colour.upper() #颜色大写
    up_colour1 = co1.upper()

    if tem_choice == 'S': #选择ZDS模板
        document = MailMerge(tem0_ZDS) #选择模板
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': str(e2.get()),
                  'engineer': str(e3.get()),
                  'wide': str(e4.get()),
                  'length': str(e5.get()),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  'colour_1':co1,
                  'element' : co1_ele,
                  'ink': ref_ink,
                  'vft': str(co1_VF_t),
                  'vfm': str(co1_VF_m),
                  'bo': str(co1_bo),
                  'bbo': str(co1_bbo),
                  'upcolour1': up_colour1,
                  'power':str(pw),
                  'inc':inc,
                  'milli': millimeter
                  }
        document.merge_pages([cust_1])
        document.write("./{}.docx".format(e1.get()))

    elif tem_choice == 'D': #选择ZDD模板
        document = MailMerge(tem1_ZDD)
        cust_1 = {'name': e1.get(),
                  'date1': '{:%B %d,%Y}'.format(date.today()),
                  'date2': '{:%Y-%m-%d}'.format(date.today()),
                  'spec_no': str(e2.get()),
                  'engineer': str(e3.get()),
                  'wide': str(e4.get()),
                  'length': str(e5.get()),
                  'n_colour': sm_colour,
                  'up_sm_colour': up_sm_colour,
                  'jixing': jixing,
                  'colour_1':co1,
                  'element' : co1_ele,
                  'ink': ref_ink,
                  'vft': str(co1_VF_t),
                  'vfm': str(co1_VF_m),
                  'bo': str(co1_bo),
                  'bbo': str(co1_bbo),
                  'upcolour1': up_colour1
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
    global co1, co2, co3, co4, co1_ele, co1_VF_m, co1_VF_t, co1_bo, co1_bbo
    if chip_len == 1:
        co1 = choice_chip_inf[0][1]
        co1_ele = choice_chip_inf[0][2]
        co1_VF_t = choice_chip_inf[0][3]
        co1_VF_m = choice_chip_inf[0][4]
        co1_bo = choice_chip_inf[0][5]
        co1_bbo = choice_chip_inf[0][6]#晶片其他参数
    elif chip_len == 2:
        co1 = choice_chip_inf[0][1]
        co2 = choice_chip_inf[1][1]
    elif chip_len == 3:
        co1 = choice_chip_inf[0][1]
        co2 = choice_chip_inf[1][1]
        co3 = choice_chip_inf[2][1]
    elif chip_len == 4:
        co1 = choice_chip_inf[0][1]
        co2 = choice_chip_inf[1][1]
        co3 = choice_chip_inf[2][1]
        co4 = choice_chip_inf[3][1]
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