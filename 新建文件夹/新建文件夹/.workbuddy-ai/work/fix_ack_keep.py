# -*- coding: utf-8 -*-
"""致谢部分各段加 keepNext（与下段同页），使致谢整体不被拆页；不使用 pageBreakBefore"""
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
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
             'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
             'pPrChange']


def set_keep_next(el, on=True):
    pPr = el.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        el.insert(0, pPr)
    kn = pPr.find(qn('w:keepNext'))
    if on:
        if kn is None:
            kn = OxmlElement('w:keepNext')
            pPr.insert(0, kn)
        # 按 schema 重排
        kids = list(pPr)
        kids.sort(key=lambda e: PPR_ORDER.index(e.tag.split('}')[1])
                  if e.tag.split('}')[1] in PPR_ORDER else 999)
        for k in kids:
            pPr.append(k)
    else:
        if kn is not None:
            pPr.remove(kn)


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
    # 致谢块 = ack .. 文档末尾（跳过末尾的空段）
    end = len(ps) - 1
    while end > ack and not ps[end].text.strip():
        end -= 1
    n = 0
    for j in range(ack, end):          # 最后一段不加
        if not ps[j].text.strip() and j > ack + 2:
            continue
        set_keep_next(ps[j]._element, True)
        n += 1
    print(f'{nm[:18]} 致谢块 [{ack}..{end}] 加 keepNext {n} 段')
    d.save(p)
