# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
nm = '装配式施工质量管理问题及优化研究——以市政项目为例'
app = wc.gencache.EnsureDispatch('Word.Application')
app.Visible=False; app.DisplayAlerts=0
doc = app.Documents.Open(os.path.join(W,nm,nm+'.docx'), ReadOnly=True)
doc.Repaginate()
for i,p in enumerate(doc.Paragraphs):
    t = p.Range.Text.replace('\r','').replace('\x07','').strip()
    if t and ('结论' in t or '致谢' in t or '参考文献' in t):
        try: pg = p.Range.Information(3)
        except Exception: pg='?'
        print(f"  i={i} pg={pg} {t[:50]!r}")
doc.Close(False); app.Quit()
