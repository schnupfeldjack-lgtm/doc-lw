# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
def norm(s): return re.sub(r'[\s．.、，,]','', s or '')
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0
for nm in ['装配式施工质量管理问题及优化研究——以市政项目为例']:
    doc = app.Documents.Open(os.path.join(W,nm,nm+'.docx'), ReadOnly=True); doc.Repaginate()
    for i,p in enumerate(doc.Paragraphs):
        t = p.Range.Text.replace('\r','').replace('\x07','').strip()
        if not t or '\t' in t: continue
        nk = norm(t)
        if nk in ('结论','致谢','参考文献'):
            print(f"  i={i} pg={p.Range.Information(3)} raw={t!r}")
        m = re.match(r'^(\d+(?:\s*[.．]\s*\d+)*)\s*(.+)$', t)
        if m and norm(m.group(2)) in ('结论','致谢','参考文献'):
            print(f"  (num) i={i} pg={p.Range.Information(3)} raw={t!r} key={norm(m.group(1))+norm(m.group(2))}")
    doc.Close(False)
app.Quit()
