# -*- coding: utf-8 -*-
"""扫描三篇论文中残留的蓝色/红色字体（模板说明文字）"""
import re, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.shared import RGBColor

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]

def tgt(c):
    """判断是否偏蓝或偏红（非纯黑/灰/白）"""
    if c is None:
        return False
    try:
        t = c if isinstance(c, tuple) else (c[0], c[1], c[2])
    except Exception:
        return False
    r, g, b = t
    mx, mn = max(t), min(t)
    if mx - mn < 40:      # 灰阶
        return False
    return True

for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    hits = []
    # 段落
    for i, para in enumerate(d.paragraphs):
        for r in para.runs:
            c = r.font.color
            if c is None or c.rgb is None:
                continue
            rgb = c.rgb
            try:
                t = (int(str(rgb)[0:2], 16), int(str(rgb)[2:4], 16), int(str(rgb)[4:6], 16))
            except Exception:
                continue
            if tgt(t):
                hits.append(('P', i, rgb, r.text[:60]))
    # 表格
    for ti, tb in enumerate(d.tables):
        for ri, row in enumerate(tb.rows):
            for ci, cell in enumerate(row.cells):
                for para in cell.paragraphs:
                    for r in para.runs:
                        c = r.font.color
                        if c is None or c.rgb is None:
                            continue
                        rgb = c.rgb
                        try:
                            t = (int(str(rgb)[0:2], 16), int(str(rgb)[2:4], 16), int(str(rgb)[4:6], 16))
                        except Exception:
                            continue
                        if tgt(t):
                            hits.append((f'T{ti}r{ri}c{ci}', -1, rgb, r.text[:60]))
    print(f'=== {nm[:20]}  命中 {len(hits)}')
    for h in hits[:40]:
        print('   ', h[0], h[2], repr(h[3]))
    print()
