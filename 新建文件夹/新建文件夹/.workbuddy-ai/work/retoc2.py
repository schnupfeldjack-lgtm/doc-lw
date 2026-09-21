# -*- coding: utf-8 -*-
"""用 Word 真实分页重算目录页码（倒序定位页码 run，兼容页码独立成 run 的情况）"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
from docx import Document

W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def norm(s):
    return re.sub(r'[\s．.、，,]', '', s or '')

app = wc.gencache.EnsureDispatch('Word.Application')
app.Visible = False; app.DisplayAlerts = 0
store = {}
for nm in FILES:
    path = os.path.join(W, nm, nm + '.docx')
    doc = app.Documents.Open(path, ReadOnly=True)
    doc.Repaginate()
    real = {}
    for p in doc.Paragraphs:
        t = p.Range.Text.replace('\r', '').replace('\x07', '').strip()
        if not t or '\t' in t:
            continue
        try: pg = p.Range.Information(3)
        except Exception: continue
        m = re.match(r'^(\d+(?:\s*[.．]\s*\d+)*)\s*(.+)$', t)
        if m:
            real.setdefault(norm(m.group(1)) + norm(m.group(2)), pg)
        else:
            nk = norm(t)
            if nk in ('结论', '致谢', '参考文献', '附录'):
                real.setdefault(nk, pg)
    store[nm] = real
    doc.Close(False)
app.Quit()

for nm in FILES:
    path = os.path.join(W, nm, nm + '.docx')
    real = store[nm]
    d = Document(path)
    fixed = unmatched = same = 0
    for p in d.paragraphs:
        if '\t' not in p.text:
            continue
        head, _, old = p.text.rpartition('\t')
        head = head.strip()
        m = re.match(r'^(\d+(?:\s*[.．]\s*\d+)*)\s*(.*)$', head)
        key = norm(m.group(1)) + norm(m.group(2)) if m else norm(head)
        want = real.get(key)
        if want is None:
            unmatched += 1; print(f'   ?? 未匹配 {head[:24]}'); continue
        want = str(want)
        if want == old.strip():
            same += 1; continue
        # 倒序找最后一个含数字的 run，替换其结尾的数字串
        hit = False
        for r in reversed(p.runs):
            mm = re.search(r'\d+\s*$', r.text)
            if mm:
                r.text = r.text[:mm.start()] + want
                hit = True; break
        if hit:
            fixed += 1
            print(f'   修正 {head[:20]}  {old.strip()} -> {want}')
        else:
            unmatched += 1
    d.save(path)
    print(f'{nm[:24]}  修正 {fixed} 条，已正确 {same} 条，未匹配 {unmatched} 条')
