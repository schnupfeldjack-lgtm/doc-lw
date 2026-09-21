# -*- coding: utf-8 -*-
"""
1) 回退正文里模板未规定的加粗（只保留模板明文要求的「摘 要：/关键词：/Abstract：/Key words：」）
2) 正文文献角标改为上标（模板明文：『在所应用段落或句子的右上角，用方括弧进行角注』）
"""
import sys, io, re, copy
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
KEEP_BOLD_PREFIX = ('摘 要：', '摘　要：', '摘 要:', '关键词：', '关键词:',
                    'Abstract：', 'Abstract:', 'Key words：', 'Key words:')
SUP_RE = re.compile(r'\[\d+\]')


def unbold_body(para):
    """去掉正文段落中的加粗（模板未规定正文可加粗）"""
    t = para.text
    if len(t) <= 60:
        return 0
    if t.strip().startswith(KEEP_BOLD_PREFIX):
        return 0
    if re.match(r'^\[\d+\]', t.strip()):     # 文献条目
        return 0
    n = 0
    for r in para.runs:
        rpr = r._element.rPr
        if rpr is None:
            continue
        for tag in ('w:b', 'w:bCs'):
            e = rpr.find(qn(tag))
            if e is not None:
                rpr.remove(e)
                n += 1
    return n


def rebuild(para, specs):
    """specs: list of (text, sup, src_run)"""
    runs = para.runs
    if not runs:
        return
    parent = runs[0]._element.getparent()
    els = [r._element for r in runs]
    for el in els:
        parent.remove(el)
    for txt, sup, src in specs:
        el = copy.deepcopy(src._element)
        for t in el.findall(qn('w:t')):
            el.remove(t)
        t = OxmlElement('w:t')
        t.text = txt
        t.set(qn('xml:space'), 'preserve')
        rpr = el.find(qn('w:rPr'))
        if rpr is not None:
            rpr.addnext(t)
        else:
            el.insert(0, t)
        rpr = el.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            el.insert(0, rpr)
        va = rpr.find(qn('w:vertAlign'))
        if sup:
            if va is None:
                va = OxmlElement('w:vertAlign')
                rpr.append(va)
            va.set(qn('w:val'), 'superscript')
        else:
            if va is not None:
                rpr.remove(va)
        parent.append(el)


for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    nb = 0
    ns = 0
    for para in d.paragraphs:
        nb += unbold_body(para)
        t = para.text
        if re.match(r'^\[\d+\]', t.strip()):     # 跳过文献条目
            continue
        if not SUP_RE.search(t):
            continue
        specs = []
        hit = 0
        for r in para.runs:
            parts = SUP_RE.split(r.text)
            marks = SUP_RE.findall(r.text)
            seq = []
            for i, pt in enumerate(parts):
                seq.append((pt, False))
                if i < len(marks):
                    seq.append((marks[i], True))
            for txt, sup in seq:
                if txt:
                    specs.append((txt, sup, r))
                    if sup:
                        hit += 1
        if hit:
            rebuild(para, specs)
            ns += hit
    print(f'{nm[:20]} 取消正文加粗 {nb} 处；角标转上标 {ns} 个')
    d.save(p)
