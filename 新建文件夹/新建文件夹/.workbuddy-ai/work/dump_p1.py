# -*- coding: utf-8 -*-
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
nm = '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'
d = Document(f'{W}\\{nm}\\{nm}.docx')
TW = 12700.0
for i, p in enumerate(d.paragraphs):
    t = p.text
    pf = p.paragraph_format
    r = p.runs[0] if p.runs else None
    sz = r.font.size.pt if (r and r.font.size) else None
    ea = None
    if r is not None:
        try:
            rpr = r._element.rPr
            if rpr is not None and rpr.rFonts is not None:
                ea = rpr.rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
        except Exception:
            pass
    ls = pf.line_spacing
    if isinstance(ls, (int, float)) and ls and ls > 10000:
        ls = round(ls/TW, 1)
    tag = ''
    if re.match(r'^图\s*\d', t.strip()): tag = '<<FIG'
    if re.match(r'^表\s*\d', t.strip()): tag = '<<TBL'
    if '\t' in t: tag += ' TOC'
    print(f'{i:4d} sz={sz} ea={ea} b={r.font.bold if r else None} al={p.alignment} ls={ls} li={round(pf.left_indent/TW,1) if pf.left_indent else None} fi={round(pf.first_line_indent/TW,1) if pf.first_line_indent else None} {tag} >> {t[:70]}')
print('--- 内联图片数:', len(d.inline_shapes))
