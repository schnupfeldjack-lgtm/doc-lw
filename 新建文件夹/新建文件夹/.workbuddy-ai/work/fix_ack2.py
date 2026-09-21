# -*- coding: utf-8 -*-
"""
对齐模板的致谢分页逻辑：
  模板第88段(结论前)、第107段(参考文献前) 有 <w:br type=page>；
  致谢标题(第134段)前**没有**分页符，只有「（空2行）」+ 1个空段。
故：去掉致谢标题的 pageBreakBefore，改为「2个空段」，并加 keepNext 防孤行。
"""
import sys, io, copy
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
        print(f'{nm[:18]} 未找到致谢')
        continue
    el = ps[ack]._element
    pPr = el.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        el.insert(0, pPr)
    # 1) 去 pageBreakBefore
    pb = pPr.find(qn('w:pageBreakBefore'))
    if pb is not None:
        pPr.remove(pb)
        print(f'{nm[:18]} 去掉致谢 pageBreakBefore')
    # 2) 加 keepNext（防标题孤行）
    kn = pPr.find(qn('w:keepNext'))
    if kn is None:
        kn = OxmlElement('w:keepNext')
        pPr.insert(0, kn)
    # 3) 前面补 2 个空段（对应模板「空2行」）
    prev = el.getprevious()
    n_empty = 0
    e = prev
    while e is not None and n_empty < 3:
        txt = ''.join(t.text or '' for t in e.iter(qn('w:t')))
        if txt.strip():
            break
        n_empty += 1
        e = e.getprevious()
    need = 2 - n_empty
    for _ in range(max(0, need)):
        blank = copy.deepcopy(el)
        # 清空文本
        for r in blank.findall(qn('w:r')):
            blank.remove(r)
        bp = blank.find(qn('w:pPr'))
        if bp is not None:
            for tag in ('w:keepNext', 'w:pageBreakBefore', 'w:keepLines'):
                t2 = bp.find(qn(tag))
                if t2 is not None:
                    bp.remove(t2)
        el.addprevious(blank)
    print(f'{nm[:18]} 致谢前原有空段 {n_empty} 个，补 {max(0, need)} 个')
    d.save(p)
