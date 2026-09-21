# -*- coding: utf-8 -*-
"""更新页眉 PAGE/NUMPAGES 域缓存（现为模板残留的"共 4 页"）"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0
for nm in FILES:
    path = os.path.join(W, nm, nm+'.docx')
    doc = app.Documents.Open(path)
    doc.Repaginate()
    doc.Fields.Update()
    for s in doc.Sections:
        for h in (s.Headers(1), s.Headers(2), s.Headers(3)):
            try: h.Range.Fields.Update()
            except Exception: pass
        for f in (s.Footers(1), s.Footers(2), s.Footers(3)):
            try: f.Range.Fields.Update()
            except Exception: pass
    doc.Save()
    n = doc.ComputeStatistics(2)
    txt = ''
    for s in doc.Sections:
        for p in s.Headers(1).Range.Paragraphs:
            if p.Range.Text.strip(): txt = p.Range.Text.strip()[:40]; break
        if txt: break
    print(f"{nm[:20]} 页数={n} 页眉={txt!r}")
    doc.Close(False)
app.Quit()
