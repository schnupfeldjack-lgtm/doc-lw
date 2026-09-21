# -*- coding: utf-8 -*-
import sys
from docx import Document
p = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/炎黄职业技术学院毕业论文模板(1).docx"
d = Document(p)
print("=== BODY paragraphs ===")
for i, par in enumerate(d.paragraphs):
    t = par.text
    if not t.strip():
        t = "(空)"
    r = None
    for run in par.runs:
        if run.text.strip():
            r = run; break
    if r is None and par.runs: r = par.runs[0]
    info = ""
    if r is not None:
        f = r.font
        info = f"sz={f.size.pt if f.size else None} name={f.name} ea={r._element.rPr.rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia') if r._element.rPr is not None and r._element.rPr.rFonts is not None else None} b={f.bold} color={f.color.rgb if f.color and f.color.type is not None else None}"
    pf = par.paragraph_format
    print(f"[{i:3d}] {t[:60]!r} | {info} | jc={pf.alignment} first={pf.first_line_indent} line={pf.line_spacing} before={pf.space_before} after={pf.space_after}")
