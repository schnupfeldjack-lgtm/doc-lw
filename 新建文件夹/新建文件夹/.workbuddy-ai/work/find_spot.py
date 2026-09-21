# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
JOBS = [
 ('BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例', 6, 8, 7),
 ('BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例', 11, 13, 12),
 ('装配式施工质量管理问题及优化研究——以市政项目为例', 9, None, 10),
]
for nm, a, b, need in JOBS:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 80); print(f'{nm[:20]}  需在 [{a}] 与 [{b}] 之间插入 [{need}]')
    ia = ib = None
    for i, p in enumerate(ps):
        t = p.text
        if re.match(r'^\[\d+\]', t.strip()) or '\t' in t:
            continue
        if f'[{a}]' in t and ia is None: ia = i
        if b and f'[{b}]' in t and ib is None: ib = i
    lo = ia if ia is not None else 0
    hi = ib if ib is not None else len(ps) - 1
    print(f'  [{a}] 首现于段{ia}，[{b}] 首现于段{ib}，候选区间 {lo}~{hi}')
    shown = 0
    for i in range(lo, min(hi + 3, len(ps))):
        t = ps[i].text.strip()
        if len(t) > 90 and not re.match(r'^\[\d+\]', t) and '\t' not in ps[i].text:
            print(f'   段{i}: {t[:110]}')
            shown += 1
            if shown >= 7: break
