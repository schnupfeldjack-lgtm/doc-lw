# -*- coding: utf-8 -*-
from pathlib import Path
from docx import Document
import re
ROOT=Path(__file__).resolve().parents[2]
DOC=ROOT/"新建文件夹/新建文件夹/装配式施工质量管理问题及优化研究——以市政项目为例/装配式施工质量管理问题及优化研究——以市政项目为例_最终优化版.docx"
d=Document(DOC)
for ti,t in enumerate(d.tables,1):
    print("\nTABLE",ti,"rows",len(t.rows),"cols",len(t.columns))
    for r in t.rows:
        print(" | ".join(c.text.replace("\n"," / ").strip() for c in r.cells))
