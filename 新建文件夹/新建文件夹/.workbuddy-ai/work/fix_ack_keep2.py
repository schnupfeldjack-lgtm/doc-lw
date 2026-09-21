# -*- coding: utf-8 -*-
"""致谢块整体不跨页：对"致  谢"标题及其后各段（除最后一段）设 keepWithNext"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def set_kn(p, on=True):
    pPr = p._element.get_or_add_pPr()
    old = pPr.find(qn('w:keepNext'))
    if old is not None: pPr.remove(old)
    if on: pPr.append(OxmlElement('w:keepNext'))

for sub in FILES:
    p_ = os.path.join(BASE, sub, sub+".docx")
    d = docx.Document(p_); ps = d.paragraphs
    idx = [i for i,p in enumerate(ps) if p.text.strip()=='致  谢']
    if not idx: print(sub[:20],"未找到致谢"); continue
    i0 = idx[-1]
    block = [i0]
    j = i0+1
    while j < len(ps):
        if ps[j].text.strip(): block.append(j)
        if len(block) >= 4: break     # 标题 + 最多3段
        j += 1
    for k in block[:-1]:
        set_kn(ps[k])
    print(f"  {sub[:20]} 致谢块段{block} 已设 keepWithNext（末段{block[-1]}除外）")
    d.save(p_)
