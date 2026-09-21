# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
nm = '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0
doc = app.Documents.Open(os.path.join(W,nm,nm+'.docx'), ReadOnly=True); doc.Repaginate()
rows=[]
for i,p in enumerate(doc.Paragraphs):
    t = p.Range.Text.replace('\r','').replace('\x07','')
    try: pg = p.Range.Information(3)
    except Exception: pg='?'
    if isinstance(pg,int) and 28 <= pg <= 32:
        pb = p.PageBreakBefore
        rows.append((pg,i,pb,repr(t[:60])))
for r in rows: print(f"pg{r[0]} i={r[1]} pbBefore={r[2]} {r[3]}")
doc.Close(False); app.Quit()
