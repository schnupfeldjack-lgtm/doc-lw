# -*- coding: utf-8 -*-
"""用 Word 导出指定页为 PDF，便于肉眼核对版式"""
import sys, os
import win32com.client as win32

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "p1"),
]
OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work"

w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    for path, tag in DOCS:
        for a, b in [(5, 7), (31, 33), (8, 9)]:
            d = w.Documents.Open(path, ReadOnly=True)
            out = os.path.join(OUT, f"{tag}_{a}-{b}.pdf")
            # Range=4 (wdExportFromTo), From=a, To=b
            d.ExportAsFixedFormat(out, 17, False, 0, 3, a, b, 0, True, True, 0, True, True, False, False)
            d.Close(False)
            print("导出", out)
finally:
    w.Quit()
