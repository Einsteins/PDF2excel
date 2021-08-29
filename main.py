import pdfminer
import pyocr
import importlib
import sys
import os
import time
import xlsxwriter
import xlwt
import xlrd
import codecs
from io import StringIO

importlib.reload(sys)
time1 = time.time()
# print("初始时间为：",time1)

import os.path
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed

#text_path = r'123.pdf'


# text_path = r'photo-words.pdf'

def parse(Filepath,Folderpath):
    '''解析PDF文本，并保存到TXT文件中'''
    Name=os.listdir(Filepath)

    for i in range(len(Name)):
        if((Name[i].find("展期协议")!=-1) or (Name[i].find("欠条协议")!=-1) or (Name[i].find("还款记录表")!=-1) or (Name[i].find("借出协议")!=-1)):
            text_path = filePath +"\\"+ Name[i]
            fp = open(text_path, 'rb')
            # 用文件对象创建一个PDF文档分析器
            parser = PDFParser(fp)
            # 创建一个PDF文档
            doc = PDFDocument()
            # 连接分析器，与文档对象
            parser.set_document(doc)
            doc.set_parser(parser)

            # 提供初始化密码，如果没有密码，就创建一个空的字符串
            doc.initialize()
            txtname = Folderpath+"\\"+ Name[i] + '.txt'
            if (os.path.exists(txtname)):
                os.remove(txtname)

            # 检测文档是否提供txt转换，不提供就忽略
            if not doc.is_extractable:

                raise PDFTextExtractionNotAllowed
            else:
                # 创建PDF，资源管理器，来共享资源
                rsrcmgr = PDFResourceManager()
                # 创建一个PDF设备对象
                laparams = LAParams()

                device = PDFPageAggregator(rsrcmgr, laparams=laparams)
                # 创建一个PDF解释其对象
                interpreter = PDFPageInterpreter(rsrcmgr, device)

                # 循环遍历列表，每次处理一个page内容
                # doc.get_pages() 获取page列表
                for page in doc.get_pages():
                    interpreter.process_page(page)
                    # 接受该页面的LTPage对象
                    layout = device.get_result()
                    # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
                    # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
                    # 想要获取文本就获得对象的text属性，
                    for x in layout:
                        if (isinstance(x, LTTextBoxHorizontal)):
                            #txtname ="D:\\python_project\\Matrerialtxt"+"\\"+Name[i]+'.txt'
                            #if(os.path.exists(txtname)):
                             #   os.remove(txtname)

                            with open(txtname, 'a') as f:
                                results = x.get_text()
                                #print(results)
                                #f.write(results + "\n")
                                try:
                                    f.write(results+"\n")
                                except UnicodeEncodeError:
                                    a=0
        #else:
        #   print("no interesting")
def postprocess(txt):
    txtname = os.listdir(txt)
    debt_n=0
    for i in range(len(txtname)):
        if(txtname[i].find("欠条协议") != -1):
            debt_n=debt_n+1
            print("找到"+str(debt_n)+"个欠条协议")
            f = open(txt+"\\"+txtname[i],'r',encoding="gbk")
            line = f.readlines()
            for n in range(len(line)):
                if (line[n].find("协议编号")!=-1):
                    #print(line[i].find("协议编号"))
                    contract_number=line[n][5:]
                    excel_table.write(1, 0, contract_number)

                    #print("欠款协议：",contract_number)
                if(line[n].find(" 欠款本金金额 人民币")!=-1):
                    m=line[n].find("，大写")

                    debt =line[n][11:m-1]
                    #print("本金",debt)
                    excel_table.write(1, 1, debt)
                if (line[n].find("年化") == 0):
                    interest = line[n][2:8]
                    excel_table.write(1, 5, interest)
                    #print("利息",interest)
                if (line[n].find("（以下简称“还款日”）") != -1):
                    form_date=line[n][1:12]
                    end_date=line[n][13:24]

                    excel_table.write(1, 2, form_date)
                    excel_table.write(1, 3, end_date)
                    #print("还款日从", form_date)
                    #print("至", end_date)

    delay_number = 0
    for j in range(len(txtname)):

        if(txtname[j].find("展期协议") != -1):
            delay_number=delay_number+1
            print("找到" + str(delay_number) + "个展期协议")
            f = open(txt + "\\" + txtname[j], 'r', encoding="gbk")
            line = f.readlines()
            for k in range(len(line)):
                if (line[k].find("协议编号") != -1):
                    #delay_contract_number = line[k][5:24]
                    delay_contract_number = line[k][5:]
                    excel_table.write(delay_number+1, 0, "展期"+delay_contract_number)
                    #print("展期协议：",delay_contract_number)
                if (line[k].find(" 欠款展期本息金额 人民币") != -1):
                    label_kk=line[k].find("，大写")

                    delay_debt1 = line[k][13:label_kk]
                    excel_table.write(delay_number + 1, 1 ,delay_debt1)
                    #print("展期欠款：",delay_debt)
                if(line[k].find("借款展期的本金金额为人民币【") !=-1):
                    label_m = line[k].find("【")
                    label_k = line[k].find("元")
                    delay_debt = line[k][label_m+1:label_k-1]
                    #print(delay_debt)
                    excel_table.write(delay_number + 1, 1, delay_debt)

                if (line[k].find("日开始按此利率计息，展") != -1):
                    label_f = line[k].find("开始按此利")
                    from_date=line[k][20:label_f]

                    label_e = line[k+1].find("后的到期")
                    end_date=line[k+1][8:20]
                    excel_table.write(delay_number + 1, 2, from_date)
                    excel_table.write(delay_number + 1, 3, end_date)
                    #print("还款日从", from_date)
                    #print("到", end_date)
                if(line[k].find("2.借款展期后的到期日为")!=-1):
                    label_3=line[k].find("日为【")
                    label_4=line[k].find("始按此利率计息")
                    from_date1=line[k][label_3+2:label_4-1]
                    from_date1.replace("【","")
                    from_date1.replace("】", "")
                    excel_table.write(delay_number + 1, 2, from_date1)
                if(line[k].find("借款展期后的到期日为【")!=-1):
                    label_5 = line[k].find("期日为【")
                    label_6 = line[k].find("日。")
                    end_date1 =line[k][label_5+3:label_6+1]
                    end_date1.replace("【", "")
                    end_date1.replace("】", "")
                    excel_table.write(delay_number + 1, 3, end_date1)



                if (line[k].find("欠款展期后的利率 年化") != -1):
                    delay_interest=line[k][12:18]
                    excel_table.write(delay_number + 1, 5, delay_interest)
                    #print("展期利息", delay_interest)
                if(line[k].find("借款展期后的借款利率为年化【") !=-1):
                    label_1 = line[k].find("年化【")
                    label_2 = line[k].find("】%")
                    delay_interest1 = line[k][label_1+3:label_2]
                    excel_table.write(delay_number + 1, 5, delay_interest1+"%")







if __name__ == '__main__':
    #material = 'D:\\1'
    #txtpath = 'D:\\1txt'
    #print("Author: bin.zhou918@gmail.com")
    print("脚本可以实现对于展期协议和欠条协议的部分数据抓取，但请注意部分数据可能由于源文件格式不统一的问题会有错误，慎用！")
    print("1:是       2：否")
    use_or_not=input("您希望使用吗？")

    c=int(use_or_not)
    #c=1
    if(use_or_not.find("1")!=-1):
    #if (c==1):
        material= input("请输入材料包的路径")
        txtpath=input("请输入希望生成的文件路径")
        foldername = os.listdir(material)

        for name in foldername:

            workbook = xlwt.Workbook(encoding='utf-8')
            excel_table  = workbook.add_sheet('sheet1')
            table_title_list = ['协议编号', '本金', '借出时间', '应还时间', '计息周期', '约定利率', '法定利率', '实际利率', '期内利息']
            exe_a = 0
            exe_b = 0
            exe_c = 1
            for i in table_title_list:
                excel_table.write(0,exe_a,i)
                exe_a+=1
            filePath=material+"\\"+name
            folderpath = txtpath+"\\"+name
            foler = os.path.exists(folderpath)
            if not foler:
                os.makedirs(folderpath)
            parse(filePath,folderpath)
            postprocess(folderpath)
            excelname =name+'.xls'
            workbook.save(folderpath + "\\"+excelname)
            time2 = time.time()
            print(name+" 完成")
            time.sleep(2)
        print("运行结束")
    else:
        print("告辞！")

    #print("Author: bin.zhou918@gmail.com")
    time.sleep(3)
    print("Author: bin.zhou918@gmail.com")
    print("Copyright (c) 2021 bin.zhou918@gmail.com All rights reserved.")
    time.sleep(0.1)