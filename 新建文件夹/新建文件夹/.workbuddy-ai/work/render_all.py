# -*- coding: utf-8 -*-
"""模板 vs 论文一并排渲染"""
import os, win32com.client as win32, pymupdf

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\pages"
os.makedirs(OUT, exist_ok=True)
TPL = BASE + r"\炎黄职业技术学院毕业论文模板(1).docx"
GEN = BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"

w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
for name, path in [("tpl", TPL), ("gen", GEN)]:
    pdf = os.path.join(OUT, f"{name}.pdf")
    d = w.Documents.Open(path, ReadOnly=True)
    d.ExportAsFixedFormat(pdf, 17, OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
                          IncludeDocProps=True, CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True)
    d.Close(False)
w.Quit()

for name in ["tpl", "gen"]:
    doc = pymupdf.open(os.path.join(OUT, f"{name}.pdf"))
    n = len(doc)
    print(f"{name}: {n} 页")
    for start in range(0, n, 8):
        pages = doc[start:start + 8]
        rows = (len(pages) + 3) // 4
        W_, H_ = 380, 538
        sheet = pymupdf.open()
        big = sheet.new_page(width=W_ * 4, height=H_ * rows)
        for i, pg in enumerate(pages):
            r, c = divmod(i, 4)
            big.show_pdf_page(pymupdf.Rect(c * W_, r * H_, (c + 1) * W_, (r + 1) * H_), doc, pg.number)
        f = os.path.join(OUT, f"{name}_sheet_{start//4+1}.png")
        big.get_pixmap(dpi=50).save(f)
        print("  ", f)
    doc.close()