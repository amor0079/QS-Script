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
        'ELPCB线路板材料核准单',
        'ELPF生产流程表'
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


excel_path = r'.\ID_table.xlsx'
def read_excel():
    workbook = xlrd.open_workbook(excel_path)
    sheet1 =  sheet1= workbook.sheet_by_name(u'Sheet1')
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    global Pro_name,tjp_unm,cus_num,ref_name,ref_ink,ref_sup,ref_texture,pcb_name,\
           pcb_sup,pin_name,pin_head,pin_head,film_name,film_sup,chip_name,easy_tape,\
           easy_sup,in_pack,out_pack
    tjp_unm = sheet1.cell(1, 1).value
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

read_excel()
#将选择添加进一个列表
a = []
#定义主逻辑
def show():
    for var in v:
        a.append(var.get())

    #模板文档
    template0_ELB = r".\tem\ELB_template.docx"
    template1_ELCK = r".\ELCK_template.docx"
    template2_ELF = r".\tem\ELF_template.docx"
    template3_ELM = r".\tem\ELM_template.docx"
    # #     # template4_ELN =
    # #     # template5_ELP =
    # #     # template6_ELQC =
    # #     # template7_ELW =
    # #     # template8_ELHC =
    # #     # template9_ELSMT =
    # #     # template10_ELT =
    # #     # template11_ELY =
    # #     # template12_FIL_H =
    # #     # template13_FIL_X =
    # #     # template14_REF =
    # #     # template15_PCB =
    # #     # template16_ELPF =
    #
    #主判断
    if a[0] ==1:
         document0 =MailMerge(template0_ELB)
         cust_1 = {'name': Pro_name,
                   'time': '{:%B %d,%Y}'.format(date.today()),
                   'pack_name': '{:%Y-%m-%d}'.format(date.today()),
                   'bianhao': '12342',
                   'custom_code': '123'
                   }
         document0.merge_pages([cust_1])
         # print(document.get_merge_fields()) 调试代码
         document0.write("./{}.docx".format(Pro_name)) #需修改
    if a[1] ==1:
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


#Button setting
Button(root, text = "确认输出文档", width = 10, command = show)\
            .pack(side='left', padx = 10, pady = 5)

Button(root, text = "退出", width = 10, command = root.quit)\
             .pack(side='right', padx = 10, pady = 5)

mainloop()
