# -*- coding: utf-8 -*-
"""严格审计：把三篇论文与模板逐类段落做 XML 级对照（pPr + 段落标记rPr + 文本run rPr）"""
import re, sys
from docx import Document
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}

TPL = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\炎黄职业技术学院毕业论文模板(1).docx"
DOCS = [
    (r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]

def xml_of(el):
    from lxml import etree
    if el is None: return "(无)"
    return etree.tostring(el, pretty_print=False).decode()

def classify(t):
    s = t.strip()
    if not s: return None
    if re.match(r'^摘\s*要[：:]', s): return '摘要'
    if re.match(r'^关键词[：:]', s): return '关键词'
    if re.match(r'^Abstract[：:]', s): return 'Abstract'
    if re.match(r'^Key\s*words[：:]', s): return 'Key words'
    if re.match(r'^目\s*录', s) and len(s) < 12: return '目录标题'
    if '\t' in s and re.search(r'\t\d+\s*$', s): return '目录条目'
    if re.match(r'^结\s*论$', s): return '结论标题'
    if re.match(r'^参\s*考\s*文\s*献$', s): return '文献标题'
    if re.match(r'^致\s*谢$', s): return '致谢标题'
    if re.match(r'^\[\d+\]', s): return '文献条目'
    if re.match(r'^图\s*\d+', s): return '图题'
    if re.match(r'^表\s*\d+', s): return '表题'
    if re.match(r'^\d+\.\d+\.\d+', s): return '三级标题'
    if re.match(r'^\d+．\d+', s): return '二级标题'
    if re.match(r'^\d+\s{1,2}\S', s): return '章标题'
    if len(s) >= 15: return '正文'
    return None

def probe(doc, name, want, limit=2):
    out = []
    n = 0
    for par in doc.paragraphs:
        c = classify(par.text)
        if c != want: continue
        n += 1
        if n > limit: continue
        pPr = par._p.find(qn('w:pPr'))
        # 去掉 rPr / sectPr 之外的东西保留
        pm_rPr = pPr.find(qn('w:rPr')) if pPr is not None else None
        truns = [r for r in par.runs if r.text.strip()]
        run = truns[0] if truns else (par.runs[0] if par.runs else None)
        run_rPr = run._element.find(qn('w:rPr')) if run is not None else None
        out.append((par.text[:38], xml_of(pPr), xml_of(pm_rPr), xml_of(run_rPr)))
    return n, out

CATS = ['章标题','二级标题','三级标题','正文','图题','表题','摘要','关键词','Abstract','Key words',
        '目录标题','目录条目','结论标题','文献标题','文献条目','致谢标题']

tpl = Document(TPL)
for cat in CATS:
    cnt, items = probe(tpl, '模板', cat, limit=1)
    print(f"\n########## {cat}  (模板命中 {cnt}) ##########")
    if not items:
        print("  模板无该母段落"); continue
    txt, pPr, pm, rr = items[0]
    print(f"  文本: {txt!r}")
    print(f"  pPr : {pPr}")
    print(f"  pmR : {pm}")
    print(f"  runR: {rr}")
    for path, nm in DOCS:
        d = Document(path)
        c2, it2 = probe(d, nm, cat, limit=1)
        print(f"  ---- {nm} (命中 {c2}) ----")
        if not it2:
            print("     缺失!"); continue
        t2, p2, m2, r2 = it2[0]
        print(f"     文本: {t2!r}")
        for label, a, b in (("pPr", pPr, p2), ("pmR", pm, m2), ("runR", rr, r2)):
            print(f"     {label} {'一致' if a==b else '★不一致'}")
            if a != b:
                print(f"        模板: {b if False else a}")
                print(f"        论文: {b}")
