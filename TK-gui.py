'''
#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/5/6 6:59
# @Author  : Amor
# @Site    : 
# @File    : TK-gui.py
# @Software: PyCharm
'''
from tkinter import *
from tkinter import messagebox
from mailmerge import MailMerge
from datetime import date
import xlrd
import os

root = Tk()
root.title('ZBOE条件票自动脚本--Design by Amor')
root.iconbitmap('./icon.ico')
group = LabelFrame(root, text = '请选择要输出的条件票', padx = 30, pady = 30)
group.pack(padx = 10, pady = 10)

# 定义checkbutton,一个元祖
LANGS = (
        'ELB包装条件票',
        'ELCK贴片加工核准单',
        'ELF封胶作业条件票',
        'ELM成品贴膜条件票',
        'ELN黏晶作业条件票',
        'ELP铆钉作业条件票',
        'ELQC中测作业条件票',
        'ELW生产规格总表',
        'ELHC后测作业条件票',
        'ELSMT SMT加工技术文件',
        'ELT成品外形条件票',
        'ELY压盖作业条件票',
        'ELFIL面膜核准单',
        'ELFIL面膜限度样品',
        'ELREF塑壳材料核准单',
        'ELPCB线路板材料核准单'
        # 'ELPF生产流程表'
        )
v = [] #定义一个列表
for long in LANGS:
    intVar = IntVar()
    v.append(intVar)
    b = Checkbutton(group,
                        text = long,
                        variable = intVar,
                        # command = choice_selection
                        )
    b.pack(anchor = W)

folder_path = r'.\{}生产条件票.format(Pro_name)'
def mkdir(folder_path):
    folder = os.path.exists(folder_path)
    if not folder:
        os.makedirs(folder_path)


excel_path = r'.\ID_table.xlsx' #程序与数据分离
def read_excel():
    workbook = xlrd.open_workbook(excel_path)
    sheet1= workbook.sheet_by_name(u'Sheet1')
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    global Pro_name,tjp_num,cus_num,ref_name,ref_ink,ref_sup,ref_texture,pcb_name,\
           pcb_sup,pin_name,pin_head,pin_head,film_name,film_sup,chip_name,easy_tape,\
           easy_sup,in_pack,out_pack
    tjp_num = sheet1.cell(1, 1).value
    Pro_name = sheet1.cell(2, 1).value
    cus_num = sheet1.cell(3, 1).value
    ref_name = sheet1.cell(4, 1).value
    ref_sup = sheet1.cell(4, 3).value
    ref_ink = sheet1.cell(5, 1).value
    ref_texture = sheet1.cell(5, 3).value
    pcb_name = sheet1.cell(6, 1).value
    pcb_sup = sheet1.cell(6, 3).value
    pin_name = sheet1.cell(7, 1).value
    pin_head = sheet1.cell(7, 3).value
    film_name = sheet1.cell(8, 1).value
    film_sup = sheet1.cell(8, 3).value
    chip_name = sheet1.cell(9, 1).value
    easy_tape = sheet1.cell(10, 1).value
    easy_sup = sheet1.cell(10, 3).value
    in_pack = sheet1.cell(11, 1).value
    out_pack = sheet1.cell(11, 1).value

# a = []
#定义主逻辑
def show():
    read_excel()
    a =[]
    for var in v:
        a.append(var.get())

    #模板文档
    template0_ELB = r".\tem\ELB_template.docx"
    template1_ELCK = r".\tem\ELCK_template.docx"
    template2_ELF = r".\tem\ELF_template.docx"
    template3_ELM = r".\tem\ELM_template.docx"
    # template4_ELN =
    # template5_ELP =
    # template6_ELQC =
    # template7_ELW =
    template8_ELHC = r".\tem\ELHC_template.docx"
    # #     # template9_ELSMT =
    # #     # template10_ELT =
    # #     # template11_ELY =
    template12_FIL_H = r".\tem\ELFIL_film_h.docx"
    template13_FIL_X = r".\tem\ELFIL_film_x.docx"



    #主判断
    if a[0] ==1: #ELB
         document0 =MailMerge(template0_ELB)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'custom_code': str(cus_num),
                   'time': '{:%Y-%m-%d}'.format(date.today()),
                   'pack_name': str(in_pack),
                   }
         document0.merge_pages([cust_1])
         document0.write("./ELB{}-{}.docx".format(tjp_num,Pro_name)) #需增加版本信息版本

    if a[1] ==1: #ELCK
         document1 =MailMerge(template1_ELCK)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'P_name':pcb_name,
                   'time': '{:%Y-%m-%d}'.format(date.today()),
                   'pack_name': str(in_pack),
                   }
         document0.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document0.write("./ELB{}-{}.docx".format(tjp_num,Pro_name)) #需修改
    if a[1] ==1: #ELCK
         document1 =MailMerge(template1_ELCK)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '',
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELCK{}-{}.docx".format(tjp_num,Pro_name)) #需修改

    if a[2] ==1: #ELF
         document1 =MailMerge(template2_ELF)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'P_name': pcb_name,
                   'R_name': ref_name,
                   'time': '{:%Y-%m-%d}'.format(date.today()),
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELF{}-{}.docx".format(tjp_num,Pro_name)) #需修改

    if a[3] ==1: #ELM
         document1 =MailMerge(template3_ELM)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'F_name': film_name,
                    'time': '{:%Y-%m-%d}'.format(date.today())
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELM{}-{}.docx".format(tjp_num,Pro_name)) #需修改

    if a[4] ==1: #ELN
         document1 =MailMerge(template4_ELN)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[5] ==1: #ELP
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[6] ==1: #ELQC
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[7] ==1: #ELW
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改

    if a[8] ==1: #ELHC
         document1 =MailMerge(template8_ELHC)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'custom_code': str(cus_num),
                   'time': '{:%Y-%m-%d}'.format(date.today())
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELHC{}-{}.docx".format(tjp_num,Pro_name)) #需修改

    if a[9] ==1: #ELSMT
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[10] ==1: #ELT
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[11] ==1: #ELY
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改

    if a[12] ==1: #ELFIL_H  缺少film_texture
         document1 =MailMerge(template12_FIL_H)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'supper': film_sup,
                   # 'texture':,
                   'time': '{:%Y-%m-%d}'.format(date.today())
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELFIL{}面膜核准单({} {}).docx".format(tjp_num, pcb_sup, Pro_name))

    if a[13] ==1: #ELFIL_X
         document1 =MailMerge(template13_FIL_X)
         cust_1 = {'name': Pro_name,
                   'bianhao': str(tjp_num),
                   'supper': film_sup,
                   'time': '{:%Y-%m-%d}'.format(date.today())
                   }
         document1.merge_pages([cust_1])
         document1.write("./ELFIL{}面膜限度样品({} {}).docx".format(tjp_num, pcb_sup, Pro_name))

    if a[14] ==1: #ELREF
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入.
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    if a[15] ==1: #ELPCB
         document1 =MailMerge(template0_ELB)
         cust_1 = {'name': '1', #需修改读入
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document1.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document1.write("./{}.docx".format(12)) #需修改
    # if a[16] ==1: #ELPF
    #      document1 =MailMerge(template0_ELB)
    #      cust_1 = {'name': '1', #需修改读入
    #                'time': '{:%B %d,%Y}'.format(date.today()),
    #                'pack_name': '{:%Y-%m-%d}'.format(date.today()),
    #                'bianhao': '12342',
    #                'custom_code': '123'
    #                }
    #      document1.merge_pages([cust_1])
    #      document1.write("./{}.docx".format(12)) #需修改


#Button setting
Button(root, text = "确认输出文档", width = 10, command = show)\
            .pack(side='left', padx = 10, pady = 5)

Button(root, text = "退出", width = 10, command = root.quit)\
             .pack(side='right', padx = 10, pady = 5)

mainloop()
