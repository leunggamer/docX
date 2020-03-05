import os
from docx import Document

docx=Document("D:\\2.docx")  #实例化一个文档对象

# print("开始读取标题")
# headflag = "Heading"
# for p in docx.paragraphs:  #遍历文档的每一段
#     if headflag in p.style.name:
#         print(p.text)
#         print(p.style.name)

print("开始读取段落")
for p in docx.paragraphs:  #遍历文档的每一段
    print("-----------------------------------------------")
    print("【段落内容】" + p.text)  #输出每一段的内容
    print("【段落样式】" + p.style.name) #output style
    print("-----------------------------------------------")

# print("统计文档样式")
# count = {}
# for p in docx.paragraphs:  #遍历文档的每一段
#     if p.style.name in count:
#         count[p.style.name] = count[p.style.name] + 1
#     else:
#         count[p.style.name] = 1
# print(count)

# print("开始读取表格>>>>>")
# tables = docx.tables #获取文件中的表格集
# table = tables[0 ]#获取文件中的第一个表格
# for i, row in enumerate(table.rows[:]):  # 读每行
#     row_content = []
# for cell in row.cells[:]:  # 读一行中的所有单元格
#         c = cell.text
#         row_content.append(c)
#     print(row_content)  # 以列表形式导出每一行数据
