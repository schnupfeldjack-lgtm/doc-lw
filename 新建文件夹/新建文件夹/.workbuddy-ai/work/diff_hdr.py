# -*- coding: utf-8 -*-
"""把模板的页眉/页脚细节导出，方便对照"""
from docx import Document
from docx.oxml.ns import qn
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
tpl = Document(BASE + r"\炎黄职业技术学院毕业论文模板(1).docx")
gen = Document(BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx")

def dump_hdr(d, nm):
    print(f"\n===== {nm} 节3 (正文) 页眉 =====")
    s = d.sections[3]
    hp = s.header
    print(f"  is_linked_to_previous = {hp.is_linked_to_previous}")
    for j, par in enumerate(hp.paragraphs):
        pPr = par._p.find(q('pPr'))
        tabs = pPr.find(q('tabs')) if pPr is not None else None
        pBdr = pPr.find(q('pBdr')) if pPr is not None else None
        bdr = ''
        if pBdr is not None:
            for b in pBdr.iter(q('bottom')) + pBdr.iter(q('top')):
                bdr += f"{b.tag.split('}')[-1]}={b.get(q('val'))}sz{b.get(q('sz'))};"
        print(f"  段{j}: 文本={par.text!r}")
        print(f"    对齐={par.alignment} tabs={tabs is not None} 边框={bdr}")
        for r in par.runs:
            rPr = r._element.find(q('rPr'))
            sz = rPr.find(q('sz')).get(q('val')) if rPr is not None and rPr.find(q('sz')) is not None else None
            szCs = rPr.find(q('szCs')).get(q('val')) if rPr is not None and rPr.find(q('szCs')) is not None else None
            rf = rPr.find(q('rFonts')) if rPr is not None else None
            asc = rf.get(q('ascii')) if rf is not None else None
            print(f"    run: 文本={r.text!r} sz={int(sz)/2 if sz else None}pt ascii={asc}")
dump_hdr(tpl, '模板')
dump_hdr(gen, '论文一')