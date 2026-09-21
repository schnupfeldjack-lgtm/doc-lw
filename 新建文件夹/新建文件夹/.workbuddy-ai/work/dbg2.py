# -*- coding: utf-8 -*-
"""列出三篇论文真正的图题/表题（排除正文句子）及其格式"""
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
        if not re.match(r'^(图|表)\s*\d+', t):
            continue
        # 真图题：形如 图3–1  + 两个空格 + 名称，且不含句号结尾的叙述
        iscap = bool(re.match(r'^(图|表)\s*\d+[–\-]\d+\s{2}\S', t)) and len(t) < 45
        pPr = p._p.find(q('pPr'))
        jc = pPr.find(q('jc')) if pPr is not None else None
        ind = pPr.find(q('ind')) if pPr is not None else None
        sp = pPr.find(q('spacing')) if pPr is not None else None
        run = None
        for r in p.runs:
            if r.text.strip(): run = r; break
        rr = run._element.find(q('rPr')) if run is not None else None
        sz = rr.find(q('sz')).get(q('val')) if rr is not None and rr.find(q('sz')) is not None else None
        rf = rr.find(q('rFonts')) if rr is not None else None
        ea = rf.get(q('eastAsia')) if rf is not None else None
        print(f"  {'题' if iscap else '文'} {t[:34]!r} sz={int(sz)/2 if sz else None} 字体={ea} "
              f"jc={jc.get(q('val')) if jc is not None else '-'} "
              f"首行缩进={(int(ind.get(q('firstLine')))/20 if ind is not None and ind.get(q('firstLine')) else 0)}pt "
              f"行距={sp.get(q('line')) if sp is not None else '-'}/{sp.get(q('lineRule')) if sp is not None else '-'}")
