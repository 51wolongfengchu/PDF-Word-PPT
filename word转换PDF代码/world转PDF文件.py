import os
from win32com.client import Dispatch

path = os.getcwd()

old_file_path = os.path.abspath('D:/14482/pythonProject/.A储存项目/江苏省肿瘤小分子靶向治疗及伴随诊断工程研究中心开放课题申报书.docx')
new_file_path = os.path.abspath('D:/14482/pythonProject/.A储存项目/江苏省肿瘤小分子靶向治疗及伴随诊断工程研究中心开放课题申报书.pdf')

word = Dispatch('Word.Application')
doc = word.Documents.Open(old_file_path)
wdFormatPDF = 17
doc.SaveAs(new_file_path, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

print('successfully')