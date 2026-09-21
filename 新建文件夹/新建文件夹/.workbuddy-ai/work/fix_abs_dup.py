# -*- coding: utf-8 -*-
"""摘要段落去重：只保留「摘 要：」+ 一段正文"""
import re
from docx import Document

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]
for path, nm in DOCS:
    d = Document(path)
    for p in d.paragraphs:
        t = p.text.strip()
        if not re.match(r'^摘\s*要[：:]', t):
            continue
        runs = p.runs
        if len(runs) <= 2:
            print(f"{nm}: 无需处理（run数={len(runs)}）")
            continue
        # 保留 run0(标签) 与 run1(正文)，删除其余
        for r in list(runs[2:]):
            r._element.getparent().remove(r._element)
        d.save(path)
        print(f"{nm}: 删除多余 run {len(runs)-2} 个，现字数={len(p.text.replace(' ',''))}")
