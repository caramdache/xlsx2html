#!/usr/bin/env python3

import xlsxwriter

with open('html2excel.py') as infile:
    exec(infile.read())

wb = xlsxwriter.Workbook('test.xlsx')
ws = wb.add_worksheet()

p = HTML2Excel(wb, ws, default_format={
    'font_name': 'Arial',
    'font_size': 10,
    'text_wrap': 1,
    'valign': 'top',
    'border': 1,
    'border_color': '#0000ff',
})    

with open('test.html') as input:
    html = input.read()
    p.feed(html)

wb.close()
