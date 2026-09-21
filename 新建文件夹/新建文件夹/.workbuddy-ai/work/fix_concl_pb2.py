# -*- coding: utf-8 -*-
"""清除结论前空行里残留的 <w:br w:type="page"/>（python-docx 读为 '\n'）"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn

BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def strip_page_br(el):
    n = 0
    for br in el.iter(qn('w:br')):
        if br.get(qn('w:type')) == 'page':
            br.getparent().remove(br); n += 1
        else:
            br.getparent().remove(br); n += 1   # 空行里的任何 br 都清掉
    for t in el.iter(qn('w:t')):
        if t.text and not t.text.strip():
            t.text = ''
    return n

for sub in FILES:
    p_ = os.path.join(BASE, sub, sub+".docx")
    d = docx.Document(p_); ps = d.paragraphs
    total = 0
    for i, p in enumerate(ps):
        if p.text.strip() == '结  论':
            for k in (i-1, i-2):
                if k >= 0:
                    n = strip_page_br(ps[k]._element)
                    if n: print(f"  段[{k}] 清除分页符/换行 {n} 个"); total += n
    # 全文兜底：任何"空段落"里都不应有 br
    for p in ps:
        if not p.text.strip():
            total += strip_page_br(p._element)
    d.save(p_)
    print(f"{sub[:24]} 清除 {total} 个")
