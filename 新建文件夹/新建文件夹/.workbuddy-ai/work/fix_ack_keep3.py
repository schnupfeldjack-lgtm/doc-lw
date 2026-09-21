# -*- coding: utf-8 -*-
"""致谢正文段落设 keepLines（段中不分页），避免末页只剩两三行"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def set_kl(p):
    pPr = p._element.get_or_add_pPr()
    if pPr.find(qn('w:keepLines')) is None:
        pPr.append(OxmlElement('w:keepLines'))

for sub in FILES:
    p_ = os.path.join(BASE, sub, sub+".docx")
    d = docx.Document(p_); ps = d.paragraphs
    idx = [i for i,p in enumerate(ps) if p.text.strip()=='致  谢']
    i0 = idx[-1]
    n=0
    for j in range(i0+1, min(i0+5, len(ps))):
        if ps[j].text.strip():
            set_kl(ps[j]); n+=1
    d.save(p_)
    print(f"{sub[:22]} 致谢正文 {n} 段设 keepLines")
