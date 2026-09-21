# -*- coding: utf-8 -*-
"""详细查论文三的表0/表1/表2/表3 全部cell内容"""
from docx import Document
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
path = BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"
d = Document(path)
for ti in [0, 1, 2, 3]:
    tb = d.tables[ti]
    print(f"\n=== 表{ti} {len(tb.rows)}x{len(tb.columns)} ===")
    seen = set()
    for ri, row in enumerate(tb.rows):
        cells = []
        for c in row.cells:
            if id(c._tc) in seen: continue
            seen.add(id(c._tc))
            cells.append(c.text.replace('\n', ' ').strip())
        print(f"  R{ri}: {cells}")