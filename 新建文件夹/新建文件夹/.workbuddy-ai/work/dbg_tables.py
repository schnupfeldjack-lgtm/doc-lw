# -*- coding: utf-8 -*-
"""检查开题报告/任务书/中期检查指导表/封面的实际填写情况"""
from docx import Document
import sys
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]
for path, nm in DOCS:
    d = Document(path)
    print(f"\n===== {nm} =====")
    # 表格 0~4：开题报告/任务书/中期检查/封面/日期小表
    for ti, tb in enumerate(d.tables[:5]):
        print(f"\n--- 表{ti} （{len(tb.rows)}行 x {len(tb.columns)}列）---")
        seen = set()
        for ri, row in enumerate(tb.rows):
            for ci, c in enumerate(row.cells):
                if id(c._tc) in seen: continue
                seen.add(id(c._tc))
                t = c.text.replace('\n', ' ').strip()
                if t:
                    print(f"  R{ri}C{ci}: {t[:80]}")