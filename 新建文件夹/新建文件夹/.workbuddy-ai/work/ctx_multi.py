# -*- coding: utf-8 -*-
import docx, re, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = [
 ("再生混凝土在装配式建筑中的应用评价——以住宅项目为例", [172,179,189]),
 ("BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例", [104,144]),
]
for sub, idxs in FILES:
    p = os.path.join(BASE, sub, sub + ".docx")
    d = docx.Document(p)
    ps = d.paragraphs
    print("="*70); print(sub, "段数", len(ps))
    for i in idxs:
        print("-"*60)
        for j in range(max(0,i-1), min(len(ps), i+2)):
            mark = ">>" if j==i else "  "
            print(f"{mark}[{j}] {ps[j].text[:300]}")
