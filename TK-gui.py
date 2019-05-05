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

root = Tk()
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

#将选择添加进一个列表
a = []
#定义主逻辑
def show():
    for var in v:
        a.append(var.get())
    print(a)
    # template0_ELB = r"C:\Users\win\Desktop\GUI_TJP_script"
    # # template1_ELCK =
    # #     # template2_ELF =
    # #     # template3_ELM =
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
    # #主判断
    # if a[0] ==1:
    #     document0 =MailMerge(template)
    #     cust_1 = {'name': 1, #需修改读入
    #               'date1': '{:%B %d,%Y}'.format(date.today()),
    #               'date2': '{:%Y-%m-%d}'.format(date.today()),
    #               # 'spec_no': e2.get(),
    #               # 'engineer': e3.get(),
    #               # 'wide': e4.get(),
    #               # 'length': e5.get(),
    #               # 'n_colour': sm_colour,
    #               # 'up_sm_colour': up_sm_colour,
    #               # 'jixing': jixing,
    #               }
    #     document0.merge_pages([cust_1])
    #     # print(document.get_merge_fields())
    #     document.write("./{}.docx".format(12)) #需修改
    #     # e1.delete(0, END)



Button(root, text = "确认输出文档", width = 10, command = show)\
            .pack(side='left', padx = 10, pady = 5)

Button(root, text = "退出", width = 10, command = root.quit)\
             .pack(side='right', padx = 10, pady = 5)

mainloop()
