import os

import xlsxwriter
from docx import Document
from xlsxwriter import Workbook

from test_protocol.table import Table

tableList = []


def init_format(workbook=None):
    if not isinstance(workbook, Workbook):
        return
    global table_title_format
    table_title_format = workbook.add_format({
        'bold': True,
        'border': 6,
        'align': 'left',  # 水平居中
        'valign': 'vcenter',  # 垂直居中
        'fg_color': '#cccccc',  # 颜色填充
        'font_name': 'Arial',
        'font_size': 10,
        'text_wrap': 1,
    })
    table_title_format.set_text_wrap(True)
    table_title_format.set_has_fill(True)

    global table_content_format1
    table_content_format1 = workbook.add_format({
        'font_name': 'Arial',
        'text_wrap': 1,
        'font_size': 10,
    })

    global table_content_format2
    table_content_format2 = workbook.add_format({
        'font_name': 'Arial',
        'color': '#633333',
        'valign': 'vcenter',  # 垂直居中
    })


def process_table_name(t: Table, cells):
    t.headName = cells[0].text
    t.tableName = t.headName.replace("Example:", "").lstrip(" \n")
    if t.headName.find("Example:") != -1:
        if t.tableName.find("+RESP:") != -1:
            separator_num1 = t.tableName.find(':')
            separator_num2 = t.tableName.find(',')
            t.tableName = t.tableName[separator_num1 + 1:separator_num2]
            return True
        elif t.tableName.find("+ACK:") != -1:
            separator_num1 = t.tableName.find('+')
            separator_num2 = t.tableName.find(',')
            t.tableName = t.tableName[separator_num1:separator_num2]
            t.tableName = t.tableName.replace(":", "-")
            return True
        else:
            return False
    else:
        return False


def process_table_table_header(t: Table, cells):
    col_title = []
    for colNum, c in enumerate(cells):
        col_title.append(c.text)
    t.colName = col_title


def process_table_body(cells, body_row):
    row = []
    for colNum, c in enumerate(cells):
        if c.text.find('-') != -1 or c.text.find('–') != -1:
            cell_content = "".join(c.text.split())
        else:
            cell_content = c.text
        if cell_content.find('\'-\'') != -1:
            cell_content = cell_content.replace("'-',", "'$-$'")
        cell_content = cell_content.replace('-', ' - ')
        cell_content = cell_content.replace('–', ' - ')
        cell_content = cell_content.replace(' - ', ' - ')
        cell_content = cell_content.replace("'$ - $'", "'-'")
        cell_content = cell_content.replace(',', ', ')
        if cell_content.startswith('\''):
            cell_content = ' ' + cell_content
        row.append(cell_content)
    body_row.append(row)


def process_docx_table(docx: Document, table_list):
    # if not isinstance(docx, Document):
    #     return
    for table in docx.tables:
        rows = table.rows
        length = len(rows)
        i = 0
        t = Table()
        body_row = []
        validate_table = True
        while i < length:
            cells = rows[i].cells
            if i == 0:
                validate_table = process_table_name(t, cells)
                if not validate_table:
                    break
            elif i == 1:
                process_table_table_header(t, cells)
            else:
                process_table_body(cells, body_row)
            i += 1
        if validate_table:
            t.body = body_row
            table_list.append(t)


# 去除名称相同的表
# def delete_repetition_table_name():


def process_docx(docx_path=None):
    if not os.path.exists(docx_path):
        print("打开的文件不存在,请检查文件路径")
        return
    # 重复表的张数
    global repetition_table
    repetition_table = []
    doc = Document(docx_path)
    table_list = []
    process_docx_table(doc, table_list)
    for t in table_list:
        print(t.tableName)


    # workbook = xlsxwriter.Workbook(docx_path)
    # init_format(workbook)


process_docx("../data/doc/GV500MA @Track Air Interface Protocol (2).docx")