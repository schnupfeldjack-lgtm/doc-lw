# -*- coding: utf-8 -*-
"""检查模板"空N行"规定在论文中的落实"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def find(ps, pred):
    return [i for i, p in enumerate(ps) if pred(p.text.strip())]

def blanks_before(ps, i):
    n = 0; j = i - 1
    while j >= 0 and not ps[j].text.strip():
        n += 1; j -= 1
    return n, (ps[j].text.strip()[:22] if j >= 0 else '(文首)')

for nm in FILES:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 78); print(nm[:30])
    checks = []
    i0 = find(ps, lambda t: t.startswith('摘 要：'))
    i1 = find(ps, lambda t: t.startswith('关键词：'))
    i2 = find(ps, lambda t: t.startswith('Abstract'))
    i3 = find(ps, lambda t: t.startswith('Key words'))
    itoc = find(ps, lambda t: t.startswith('目') and '录' in t[:4])
    itocf = find(ps, lambda t: '\t' in t or False)
    # 目录首条目
    first_toc = None
    for i, p in enumerate(ps):
        if '\t' in p.text:
            first_toc = i; break
    if i0: checks.append(('摘要前（应空1行）', i0[0], 1))
    if i1: checks.append(('摘要→关键词（应空2行）', i1[0], 2))
    if i2: checks.append(('关键词→Abstract（应空1行）', i2[0], 1))
    if i3: checks.append(('Abstract→Key words（应空2行）', i3[0], 2))
    if first_toc: checks.append(('目录标题→首条目（应空1行）', first_toc, 1))
    for label, i, need in checks:
        n, prev = blanks_before(ps, i)
        flag = 'OK' if n == need else ('NG' if n < need else 'NG(多)')
        print(f'   {flag:<6} {label:<28} 实测={n} 应为={need}   上一非空: {prev}')
