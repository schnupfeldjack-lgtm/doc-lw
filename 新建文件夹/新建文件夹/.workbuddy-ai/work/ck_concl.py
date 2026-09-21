# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0
for nm in FILES:
    doc = app.Documents.Open(os.path.join(W,nm,nm+'.docx'), ReadOnly=True); doc.Repaginate()
    print("="*66); print(nm[:26])
    cache=[]
    for p in doc.Paragraphs:
        t = p.Range.Text.replace('\r','').replace('\x07','')
        try: pg=p.Range.Information(3)
        except Exception: pg='?'
        cache.append((t, p.PageBreakBefore, pg))
    for i,(t,pb,pg) in enumerate(cache):
        if t.strip() in ('结  论','参 考 文 献','致  谢'):
            print(f"  [{i}] {t.strip()!r} pbBefore={pb} pg={pg}")
            for k in (i-1,i-2):
                if k>=0: print(f"       [{k}]={cache[k][0]!r} pb={cache[k][1]} pg={cache[k][2]}")
    doc.Close(False)
app.Quit()
