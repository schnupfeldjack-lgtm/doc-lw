# -*- coding: utf-8 -*-
"""扫描模板说明文字残留 + 分页符位置"""
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
PAT = ['注：', '正文中公式', '本页为独立页', '（空1行）', '（空2行）', '说明：', '例如：',
       '××', '餐饮', '可作为正文', '作为正文', '小四号宋体', '五号宋体', '四号宋体',
       '黑体小二', 'Times New Roman，加粗', '请仔细阅读', '毕业论文（设计）的基本框架']

for path, nm in DOCS:
    d = Document(path)
    print(f"\n===== {nm} =====")
    hit = 0
    for i, p in enumerate(d.paragraphs):
        t = p.text
        for pat in PAT:
            if pat in t:
                print(f"  ★残留 [{i}] {t[:70]!r}")
                hit += 1
                break
    if not hit: print("  无残留")
    # 表格内残留
    for ti, tb in enumerate(d.tables):
        for row in tb.rows:
            for c in row.cells:
                for pat in PAT:
                    if pat in c.text:
                        print(f"  ★表格残留 表{ti}: {c.text[:60]!r}")
                        break
    # 分页符位置
    for i, p in enumerate(d.paragraphs):
        brs = [b for b in p._p.findall(f".//{q('br')}") if b.get(q('type')) == 'page']
        if brs:
            # 判断分页符在文本 run 之前还是之后
            seq = []
            for ch in p._p.iter():
                pass
            pos = '文本之后'
            for r in p.runs:
                if r._element.findall(f".//{q('br')}"):
                    pos = '在该run内'
            print(f"  分页符 [{i}] 文本={p.text[:20]!r} ({pos})")
