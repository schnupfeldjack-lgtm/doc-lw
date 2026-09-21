# -*- coding: utf-8 -*-
"""三篇都加：致谢段前分页（让致谢独占一页，与三篇保持一致）"""
import re
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

for PATH in DOCS:
    d = Document(PATH)
    ps = [p for p in d.paragraphs if re.match(r'^致\s*谢$', p.text.strip())]
    assert len(ps) == 1, f"{PATH[-20:]}: 致谢 命中 {len(ps)}"
    p = ps[0]
    pPr = p._p.find(q('pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
    if pPr.find(q('pageBreakBefore')) is not None:
        continue
    e = OxmlElement('w:pageBreakBefore')
    st = pPr.find(q('pStyle'))
    if st is not None: st.addnext(e)
    elif pPr.find(q('keepNext')) is not None: pPr.find(q('keepNext')).addnext(e)
    else: pPr.insert(0, e)
    d.save(PATH)
    print(f"{PATH[-20:]}: 已为致谢加段前分页")
print("完成")
