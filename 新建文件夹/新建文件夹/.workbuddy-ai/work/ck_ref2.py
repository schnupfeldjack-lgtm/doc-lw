# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
for nm in ['装配式施工质量管理问题及优化研究——以市政项目为例',
           'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 78); print(nm[:30])
    for p in ps:
        t = p.text.strip()
        if re.match(r'^\[\d+\]', t):
            print('   ', t[:78])
    print('  --- 正文角注首次出现顺序 ---')
    seq, seen = [], set()
    for p in ps:
        t = p.text.strip()
        if re.match(r'^\[\d+\]', t) or '\t' in p.text:
            continue
        for m in re.finditer(r'\[(\d+)\]', p.text):
            n = int(m.group(1))
            if n not in seen:
                seen.add(n); seq.append(n)
    print('   ', seq)
    print('   是否严格递增1..N:', 'OK' if seq == list(range(1, len(seq)+1)) else 'NG')
