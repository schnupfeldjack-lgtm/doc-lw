# -*- coding: utf-8 -*-
"""【10】里剩余的"不符"正文段是哪几句"""
import re
from docx import Document
from docx.oxml.ns import qn
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]
for path, nm in DOCS:
    d = Document(path)
    print(f"\n===== {nm} =====")
    for p in d.paragraphs:
        t = p.text.strip()
        if not (len(t) >= 60 and not re.match(r'^\d', t) and '\t' not in p.text
                and not t.startswith('[') and not re.match(r'^摘\s*要[：:]', t)
                and not t.startswith('Abstract') and not t.startswith('Key words')):
            continue
        pPr = p._p.find(q('pPr'))
        sp = pPr.find(q('spacing')) if pPr is not None else None
        ind = pPr.find(q('ind')) if pPr is not None else None
        run = None
        for r in p.runs:
            if r.text.strip(): run = r; break
        rr = run._element.find(q('rPr')) if run is not None else None
        sz = rr.find(q('sz')).get(q('val')) if rr is not None and rr.find(q('sz')) is not None else None
        line = sp.get(q('line')) if sp is not None else None
        lrule = sp.get(q('lineRule')) if sp is not None else None
        fl = ind.get(q('firstLine')) if ind is not None else None
        bad = (sz != '24') or (line != '360' or lrule != 'auto') or (fl != '480')
        if bad:
            print(f"  ★ sz={int(sz)/2 if sz else None} 行距={line}/{lrule} 缩进={int(fl)/20 if fl else 0}pt :: {t[:40]!r}")
