# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc, fitz
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0
doc = app.Documents.Open(os.path.join(W,'炎黄职业技术学院毕业论文模板(1).docx'), ReadOnly=True)
pdf = os.path.join(W,'.workbuddy-ai','work','out','TPL.pdf')
doc.ExportAsFixedFormat(pdf, 17)
n = doc.ComputeStatistics(2)
doc.Close(False); app.Quit()
d = fitz.open(pdf)
print("模板页数:", n, "pdf:", len(d))
for i in range(len(d)):
    tx = d[i].get_text('text').strip()
    lines = [l for l in tx.split('\n') if l.strip()]
    print(f"\n--- p{i+1} ({len(lines)}行) ---")
    for l in lines[:8]: print("   ", l[:64])
