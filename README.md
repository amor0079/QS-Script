基于python实现简单办公自动化
 # 0.环境
操作系统:windows
python版本:3.6.5
第三方库:docx-mailmerge 0.4.0/xlrd:1.2.0
其余为内置库就不列举了
# 1.前言
中小型企业,因为自身资金实力等各方面原因并没有公司内部的集成控制软件,没有数据库.单令人扎心的是实际工作往往是有非常多的word表格需要处理.然后原本说好的设计工作变成了无脑填表游戏(还很容易犯错误,可能比较抵触,大脑自动进入休眠状态了吧)0-0!!!
这时候轮到Python登场了.
需求:减少无脑输入相同信息次数(人话:自动填写批量word表格)
# 2.效果图
![图片](https://uploader.shimo.im/f/rguAJ5rPdFIGvkq8.png!thumbnail)
# 3.第三方库的安装
为什么要将第三方库单独拎出来，因为此脚本用到了一个非常规，或者相对比较冷门的第三方库docx-mailmerge.为什么不用常规的pydocx(因为功能达不到实际要求) 额，笔者目前的理解pydocx能做的事情相对比较low，对于实际比较复杂的word表格与格式有点束手无策。
docx-mailmerge 具体安装参考[https://pbpython.com/python-word-template.html](https://pbpython.com/python-word-template.html)
其余第三方库使用简单pip 就可以了。
# 4.设计思路
![图片](https://uploader.shimo.im/f/z1JSaHKSHRsSdxku.png!thumbnail)
# 5.代码实现
Talk is cheap,show me the code!
## 5.1 excel 读取数据
```
excel_path = r'.\ID_table.xlsx' #程序与数据分离
def read_excel():
    workbook = xlrd.open_workbook(excel_path)  #打开一个excel文件
    sheet1= workbook.sheet_by_name(u'Sheet1')  #选择一个Sheet1 工作表
    # nrows = sheet1.nrows
    # ncols = sheet1.ncols
    global Pro_name,tjp_num,cus_num,ref_name,\
           ref_ink,ref_sup,ref_texture,pcb_name,\
           pcb_sup,pin_name,pin_head,pin_head,film_name,\
           film_sup,chip_name,easy_tape,\
           easy_sup,in_pack,out_pack   #全局可以用.不知道其他方式可以不.
    #读取excel表想对应的数值
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
```
## 5.2 word模板制作
这个具体可以参考:[https://pbpython.com/python-word-template.html](https://pbpython.com/python-word-template.html)
需要注意的是这个模块只适合2007以上的word 版本。
![图片](https://uploader.shimo.im/f/4yMB8ypBoHAFtEGa.png!thumbnail)
在模板文档里设置替代位置。域名可以为自己定义的变量。
![图片](https://uploader.shimo.im/f/TaSyhnjPuSQz2bUg.png!thumbnail)
像这样。![图片](https://uploader.shimo.im/f/WlSIjs9Apccrse1n.png!thumbnail)
然后将这些文档存好，记下路径。
## 5.3 Tkinter 制作简单的交互界面
使用了Tkinter 这个内置库，学习曲线比PyQt 容易些。
这边使用了LabelFrame与checkbutton.
```
root = Tk()
root.title('条件票自动脚本--Design by Amor')
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
#Button setting
Button(root, text = "确认输出文档", width = 10, command = show)\
            .pack(side='left', padx = 10, pady = 5)

Button(root, text = "退出", width = 10, command = root.quit)\
             .pack(side='right', padx = 10, pady = 5)

mainloop()
```
## 5.4 python内对于数据清洗，整理与输出
从Tkinter的checkbutton中读取选择内容。
```
def show():
    read_excel()
    new_folder()
    a =[]
    for var in v:
        a.append(var.get())

    #模板文档
    template0_ELB = r".\tem\ELB_template.docx"
    template1_ELCK = r".\tem\ELCK_template.docx"
    template2_ELF = r".\tem\ELF_template.docx"
    #此处省略
#主判断
if a[0] ==1: #ELB
     document0 =MailMerge(template0_ELB) #选择模板
     cust_1 = {'name': Pro_name,
               'bianhao': str(tjp_num),
               'custom_code': str(cus_num),
               'time': '{:%Y-%m-%d}'.format(date.today()),
               'pack_name': str(in_pack),
               }
     document0.merge_pages([cust_1])
     document0.write("./{}/ELB{}-{}.docx".format(new_folder_name, tjp_num,Pro_name)) #使用format重命名文件名
```
# 6.总结
本文是笔者根据实际工作需求而制作的脚本,代码相对粗糙,代码不足或者有更好的想法，请多指点。全文的最核心的其实是docx-mailmerge 库。为了实现目前这种word操作，笔者花了不少心思与时间，算是为了实际需求Python小入门了。
如果你喜欢乐高，那么来用Python吧。
本文涉及的完整代码Github链接：[https://github.com/amor0079/QS-Script/blob/master/TK-gui.py](https://github.com/amor0079/QS-Script/blob/master/TK-gui.py)
