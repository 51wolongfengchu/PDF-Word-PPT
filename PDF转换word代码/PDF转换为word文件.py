from pdf2docx import Converter

cv = Converter("1.docx")

print(cv)

cv.convert("1.pdf", start=0, end=None)
cv.close()

print("转换完成")