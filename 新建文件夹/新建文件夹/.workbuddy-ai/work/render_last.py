# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
import fitz

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'
os.makedirs(OUT, exist_ok=True)
FILES = [
    ('P1', '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'),
    ('P2', '装配式施工质量管理问题及优化研究——以市政项目为例'),
    ('P3', 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'),
]
app = wc.gencache.EnsureDispatch('Word.Application')
app.Visible = False
app.DisplayAlerts = 0
for tag, nm in FILES:
    path = f'{W}\\{nm}\\{nm}.docx'
    pdf = f'{OUT}\\{tag}.pdf'
    doc = app.Documents.Open(path, ReadOnly=True)
    doc.ExportAsFixedFormat(pdf, 17)
    n = doc.ComputeStatistics(2)
    doc.Close(False)
    d = fitz.open(pdf)
    print(f'{tag} 页数={n} pdf={len(d)}')
    for i in range(len(d)):
        pg = d[i]
        tx = pg.get_text('text').strip()
        lines = [l for l in tx.split('\n') if l.strip()]
        last = lines[-1][:40] if lines else ''
        print(f'   p{i+1:2d} 行数={len(lines):3d} 末行: {last}')
    d.close()
app.Quit()
