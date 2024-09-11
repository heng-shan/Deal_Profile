import numpy as np
import pandas as pd
from win32com.client import Dispatch,DispatchEx
import win32com
import win32com.client
import os
from docx import Document
import PyPDF2
import psutil
from docx2pdf import convert
import threading
from shutil import copyfile
import time
import re
src="D:\\研究生\\9.12"

def printPids():
    pids = psutil.pids()
    for pid in pids:
        try:
            p = psutil.Process(pid)
            # print('pid=%s,pname=%s' % (pid, p.name()))
            # 关闭excel进程
            if p.name() == 'wps.exe':
                cmd = 'taskkill /F /IM wps.exe'
                os.system(cmd)
        except Exception as e:
            # continue
            print(e)



file=os.listdir(src)

print(file)
if "temp.pdf"  not in file:
    convert(src+"\\temp.docx",src+"\\temp.pdf")
else:
    print("temp.pdf已存在！")


excel_path=""
for i in range(len(file)):
    if file[i][-2]=="s" and file[i][-5]==".":
        excel_path=src+"\\"+file[i]
        break

student = pd.read_excel(excel_path)

teacher=""

pdf_path=src+"\\temp.pdf"
pdf=open(pdf_path,'rb')
pdfreader=PyPDF2.PdfReader(pdf)
v=np.zeros(len(pdfreader.pages))

save_teacher = "分类"
if "分类" not in file:
    os.chdir(src)
    os.makedirs(save_teacher)
else:
    print("该文件夹已存在！")

save_teacher_path=src+"\\"+save_teacher


all_teacher=[]
for i in range(student.shape[0]):
    if student['*面试老师'][i] not in all_teacher:
        all_teacher.append(student['*面试老师'][i])
count=0
dic={}
dic_end={}
student_name=[]
student_page=[]
for page in range(len(pdfreader.pages)):
    pageobj = pdfreader.pages[page]
    text = pageobj.extract_text()
    if "姓名" in text:
        name=re.findall("姓名(.*)政治面貌",text)
        if name:
            dic[name[0].strip()] = page + 1
        else:
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!",page)


t1=time.time()

save_teacher_path=src+"\\"+save_teacher
outputFile=""


teacher="hh"


app = win32com.client.Dispatch('kwps.Application')
app.Visible = 1
# 打开word，经测试要是绝对路径
doc = app.Documents.Open(src + "\\temp" + ".docx")
word = win32com.client.DispatchEx('kwps.Application')
word.Visible = 1
doc1 =0


for i in range(student.shape[0]):
    if student['*面试老师'][i]==teacher:
        for page in range(dic[student['*姓名'][i]], len(pdfreader.pages)):
        # for page in range(dic[student['*姓名'][i]],len(pdfreader.pages)):
            pageobj = pdfreader.pages[page]
            text = pageobj.extract_text()
            if "韩珂" in text and "审核人" in text:
                end=page
                break
        if end!=0:
            print(student['*姓名'][i],dic[student['*姓名'][i]],end+1)
            objRectangles = doc.ActiveWindow.Panes(1).Pages(dic[student['*姓名'][i]])
            doc.Application.ActiveDocument.Range().GoTo(1, 1, dic[student['*姓名'][i]]).Select()
            start = app.Selection.Start.numerator
            doc.Application.ActiveDocument.Range().GoTo(1, 1, end + 2).Select()
            app.Selection.MoveLeft()
            # time.sleep(1)
            end = app.Selection.Start.numerator
            doc.Range(start, end).Select()
            app.ActiveDocument.ActiveWindow.Selection.Copy()
            s = word.Selection
            s.MoveRight(1, doc1.Content.End)  # 将光标移动到文末
            word.Selection.InsertBreak(1)
            s.Paste()

    else:
        if doc1!=0:
            doc1.Save()
        print(teacher,"已经处理完毕！！！！")
        teacher=student['*面试老师'][i]
        outputFile = save_teacher_path + "\\" + teacher + ".docx"
        new_word = win32com.client.DispatchEx('kwps.Application')
        new_doc = new_word.Documents.Add()
        new_doc.SaveAs(outputFile)
        new_doc.Close()
        doc1 = word.Documents.Open(outputFile)
        end=0
        for page in range(dic[student['*姓名'][i]],len(pdfreader.pages)):
            pageobj = pdfreader.pages[page]
            text = pageobj.extract_text()
            if "韩珂" in text and "审核人" in text:
                end=page
                break
        if end!=0:
            objRectangles = doc.ActiveWindow.Panes(1).Pages(dic[student['*姓名'][i]])
            doc.Application.ActiveDocument.Range().GoTo(1, 1, dic[student['*姓名'][i]]).Select()
            start = app.Selection.Start.numerator
            doc.Application.ActiveDocument.Range().GoTo(1, 1, end + 2).Select()
            # 往左移一下
            app.Selection.MoveLeft()
            end = app.Selection.Start.numerator
            doc.Range(start, end).Select()
            app.ActiveDocument.ActiveWindow.Selection.Copy()
            s = word.Selection
            s.MoveRight(1, doc1.Content.End)  # 将光标移动到文末
            s.Paste()

    print(i+1,".  ",student['*姓名'][i],"处理完毕！")


t2=time.time()
print("耗时:   ",t2-t1)

doc1.Close()
doc.Close()
printPids()








