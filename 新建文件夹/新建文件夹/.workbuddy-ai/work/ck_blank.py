# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
def probe(ps, idx, label):
    n = 0
    j = idx - 1
    while j >= 0 and not ps[j].text.strip():
        n += 1; j -= 1
    prev = ps[j].text.strip()[:26] if j >= 0 else '(文首)'
    print(f'   {label} 前空行数={n}  上一非空段: {prev}')
for nm in FILES:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 70); print(nm[:28])
    for i, p in enumerate(ps):
        t = p.text.strip()
        if t == '结  论': probe(ps, i, '结  论')
        if t == '参 考 文 献': probe(ps, i, '参 考 文 献')
        if t == '致  谢': probe(ps, i, '致  谢')
