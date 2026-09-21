# -*- coding: utf-8 -*-
"""把章标题「1 绪论」从目录节（无页眉）挪到正文节起始（带页眉），避免视觉上孤悬且无页眉
做法：把当前挂在章标题段落上的目录节 sectPr 移到目录最后一条（参考文献），章标题自然进入正文节
"""
import re, copy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
    BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
    BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx",
]

for path in DOCS:
    d = Document(path)
    paras = d.paragraphs
    # 定位目录标题、最后 TOC 条目、第一章标题
    toc_idx = chap1_idx = None
    last_toc_idx = None
    for i, p in enumerate(paras):
        t = p.text.strip()
        if toc_idx is None and re.match(r'^目\s*录$', t):
            toc_idx = i
        if re.match(r'^\d+\s{1,2}\S', t) and '\t' not in t and len(t) < 40 and chap1_idx is None:
            chap1_idx = i
    # 最后一个目录条目（参考文献）；可能带 "\t<页码>" 后缀
    for i in range(toc_idx, chap1_idx):
        if paras[i].text.strip().startswith('参考文献'):
            last_toc_idx = i
    assert toc_idx and chap1_idx and last_toc_idx, (toc_idx, chap1_idx, last_toc_idx)

    # 取出章标题上的 sectPr（目录节设置）
    chap1_pPr = paras[chap1_idx]._p.find(q('pPr'))
    chap1_sect = chap1_pPr.find(q('sectPr')) if chap1_pPr is not None else None
    assert chap1_sect is not None
    new_sect = copy.deepcopy(chap1_sect)
    chap1_pPr.remove(chap1_sect)
    # 放到最后一条 TOC 上
    last_pPr = paras[last_toc_idx]._p.find(q('pPr'))
    if last_pPr is None:
        last_pPr = OxmlElement('w:pPr'); paras[last_toc_idx]._p.insert(0, last_pPr)
    old = last_pPr.find(q('sectPr'))
    if old is not None:
        last_pPr.remove(old)
    last_pm = last_pPr.find(q('rPr'))
    if last_pm is not None:
        last_pm.addnext(new_sect)
    else:
        last_pPr.append(new_sect)
    d.save(path)
    print(f"{path[-26:]}: 目录节结束处移至段落#{last_toc_idx}(参考文献)；章标题#{chap1_idx} 转入正文节")
print("完成")