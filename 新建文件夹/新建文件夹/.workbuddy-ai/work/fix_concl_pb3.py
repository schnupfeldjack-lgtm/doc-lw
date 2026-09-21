# -*- coding: utf-8 -*-
"""结论独立页：把段前分页从"结  论"标题移到它前面第 2 个空段上，
   使新页顶部为：空行、空行、结  论（避免空行独占一页造成空白页）"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

def set_pb(p, on):
    pPr = p._element.get_or_add_pPr()
    old = pPr.find(qn('w:pageBreakBefore'))
    if old is not None:
        pPr.remove(old)
    if on:
        pPr.append(OxmlElement('w:pageBreakBefore'))

for sub in FILES:
    p_ = os.path.join(BASE, sub, sub+".docx")
    d = docx.Document(p_); ps = d.paragraphs
    for i, p in enumerate(ps):
        if p.text.strip() == '结  论':
            set_pb(p, False)                      # 结论标题不再单独分页
            # 往前数 2 个空段，给靠前的那个设分页
            blanks = []
            j = i - 1
            while j >= 0 and not ps[j].text.strip() and len(blanks) < 2:
                blanks.append(j); j -= 1
            if len(blanks) == 2:
                set_pb(ps[blanks[1]], True)       # blanks[0]=i-1, blanks[1]=i-2
                print(f"  结论段[{i}]：分页改设于空段[{blanks[1]}]")
            else:
                print(f"  !! 结论前空段不足: {blanks}")
    d.save(p_)
    print(f"{sub[:24]} 已保存")
