# -*- coding: utf-8 -*-
import docx, os
from docx.oxml.ns import qn
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
T = [("再生混凝土在装配式建筑中的应用评价——以住宅项目为例",[172,179,189]),
     ("BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例",[104,144])]
for sub, idxs in T:
    d = docx.Document(os.path.join(BASE, sub, sub+".docx"))
    ps=d.paragraphs
    print("="*70); print(sub)
    for i in idxs:
        print(f"-- 段{i} run数={len(ps[i].runs)}")
        for k,r in enumerate(ps[i].runs):
            va = r._element.find(qn('w:rPr'))
            sup=False
            if va is not None:
                v=va.find(qn('w:vertAlign'))
                if v is not None: sup = v.get(qn('w:val'))=='superscript'
            print(f"   r{k} sup={sup} {r.text!r}")
