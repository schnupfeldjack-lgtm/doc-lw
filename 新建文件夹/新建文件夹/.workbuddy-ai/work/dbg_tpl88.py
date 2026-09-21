# -*- coding: utf-8 -*-
import sys, io, docx
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
d = docx.Document(r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/炎黄职业技术学院毕业论文模板(1).docx")
ps = d.paragraphs
for i in range(80, 100):
    if i < len(ps): print(f"TPL[{i}] {ps[i].text[:70]!r}")
