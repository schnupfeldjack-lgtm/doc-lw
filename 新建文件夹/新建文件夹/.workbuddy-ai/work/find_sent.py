# -*- coding: utf-8 -*-
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
QUERY = [
 ('装配式施工质量管理问题及优化研究——以市政项目为例', ['装配率', '评价标准', '信息化', 'BIM', '数据']),
 ('BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例', ['质量', '施工技术', '预制']),
]
for nm, kws in QUERY:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 80); print(nm[:24])
    start = 0
    for i, p in enumerate(ps):
        if p.text.strip() == '结  论': start = i
        break
    for i, p in enumerate(ps):
        t = p.text.strip()
        if not t or '\t' in p.text or re.match(r'^\[\d+\]', t):
            continue
        if i < 90 or len(t) < 60:
            continue
        for kw in kws:
            if kw in t:
                print(f'  段{i} <{kw}>: {t[:150]}')
                break
        if i > 240: break
