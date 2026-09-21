# -*- coding: utf-8 -*-
"""恢复致谢独立页（与模板 p10 参考文献 / p11 致谢 的呈现一致），并清理多余空段"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]

for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    ps = d.paragraphs
    ack = None
    for i, q in enumerate(ps):
        if q.text.strip() == '致  谢':
            ack = i
            break
    if ack is None:
        continue
    el = ps[ack]._element
    pPr = el.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        el.insert(0, pPr)
    # 加 pageBreakBefore
    if pPr.find(qn('w:pageBreakBefore')) is None:
        pb = OxmlElement('w:pageBreakBefore')
        pPr.insert(0, pb)
    # 保留 keepNext
    if pPr.find(qn('w:keepNext')) is None:
        kn = OxmlElement('w:keepNext')
        pPr.insert(0, kn)
    # 致谢标题前的空段精简为 1 个（对应模板「空2行」）
    blanks = []
    e = el.getprevious()
    while e is not None:
        txt = ''.join(t.text or '' for t in e.iter(qn('w:t')))
        if txt.strip():
            break
        blanks.append(e)
        e = e.getprevious()
    for b in blanks[1:]:        # 只保留紧邻的 1 个
        b.getparent().remove(b)
    print(f'{nm[:18]} 致谢段={ack} 前置空段 {len(blanks)} -> 1')
    d.save(p)
