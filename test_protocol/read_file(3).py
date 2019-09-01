import xlrd
import xlsxwriter
from docx import Document

from table import Table
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

#获取打开文件的路径
file_path = filedialog.askopenfilename()
doc = Document(file_path)

#设置保存文件的路径
if doc != -1:
    save_path = filedialog.asksaveasfilename(title=u'保存文件', filetypes=[("Excel工作簿", ".xlsx")], defaultextension=".xlsx")
    workbook = xlsxwriter.Workbook(save_path)
tableList = []

table_title_format = workbook.add_format({
    'bold': True,
    'border': 6,
    'align': 'left',  # 水平居中
    'valign': 'vcenter',  # 垂直居中
    'fg_color': '#cccccc',  # 颜色填充
    'font_name': 'Times New Roman',
    'font_size': 10,
    'text_wrap': 1,
})

table_content_format1 = workbook.add_format({
    'font_name': 'Times New Roman',
    'text_wrap': 1,
    'font_size': 10,
})
table_content_format2 = workbook.add_format({
    'font_name': 'Calibri',
    'color': '#633333',
    'valign': 'vcenter',  # 垂直居中
})

# 重复表的张数
repet_table = []

for table in doc.tables:
    rows = table.rows
    length = len(rows)
    i = 0
    t = Table()
    listRow = []
    validateTable = True
    while i < length:
        cells = rows[i].cells
        if i == 0:
            tn = cells[0].text
            t.headName = cells[0].text
            # originalName = t.headName
            t.tableName = t.headName.replace("Example:", "").lstrip(" \n")
            if t.headName.find("Example:") == -1 or t.tableName.find("+RESP:") == -1:
                validateTable = False
                break
            separator_Num1 = t.tableName.find(':')
            separator_Num2 = t.tableName.find(',')
            t.tableName = t.tableName[separator_Num1+1:separator_Num2]

            # t.tableName = t.tableName.replace(":", "-")
            try:
                worksheet = workbook.add_worksheet(t.tableName)
                worksheet.merge_range('A1:D1', t.headName, table_title_format)
            except Exception:
                repet_table.append(t.tableName + "2")
                break
        elif i == 1:
            colTitle = []
            for colNum, c in enumerate(cells):
                colTitle.append(c.text)
                worksheet.write(i, colNum, c.text)
                # worksheet.set_row(i,  table_content_format2)
            t.colName = colTitle
            # for colNum, c in enumerate(t.colName):
            #     worksheet.write(i, colNum, c)

        else:
            row = []
            for colNum, c in enumerate(cells):
                # c.text = c.text.strip("-").replace("-", " - ")
                if c.text.find('-') != -1 or c.text.find('–') != -1:
                    contentW1 = c.text.split()
                    contentW = "".join(contentW1)
                    #contentW = contentW.join(contentW.split())
                else:
                    contentW = c.text
                if contentW.find('\'-\'') != -1:
                    contentW = contentW.replace("'-',", "'$-$'")
                contentW = contentW.replace('-', ' - ')
                contentW = contentW.replace('–', ' - ')
                contentW = contentW.replace(' - ', ' - ')
                contentW = contentW.replace("'$ - $'", "'-'")
                contentW = contentW.replace(',', ', ')
                row.append(contentW)
                # content = c.text if not c.text.startswith('\'') else '\''+c.text
                worksheet.write(i, colNum, contentW)
                worksheet.set_row(i, 25, table_content_format1)
            listRow.append(row)
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 15)
            worksheet.set_column('C:C', 35)
            worksheet.set_column('D:D', 15)
        i += 1
    if validateTable:
        t.body = listRow
        # t.tostring()
        tableList.append(t)
workbook.close()


print("一共处理 " + str(len(tableList)) + " 张表")
print("一共有 " + str(len(repet_table)) + " 张表产生冲突")
for name in repet_table:
    print(name)
input('请按Enter键结束')


