# -*- coding: utf-8 -*-
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
        if re.match(r'^摘\s*要[：:]', t) or t.startswith('关键词'):
            print(f"\n  段落: {t[:40]}…  全文长度={len(t)} 去重空格后={len(t.replace(' ',''))}")
            for r in p.runs:
                rr = r._element.find(q('rPr'))
                b = rr is not None and rr.find(q('b')) is not None
                sz = rr.find(q('sz')).get(q('val')) if rr is not None and rr.find(q('sz')) is not None else None
                rf = rr.find(q('rFonts')) if rr is not None else None
                ea = rf.get(q('eastAsia')) if rf is not None else None
                print(f"      run b={b} sz={int(sz)/2 if sz else None} ea={ea} text={r.text[:38]!r}")
