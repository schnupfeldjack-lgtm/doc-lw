# -*- coding: utf-8 -*-
"""全量导出 PDF 并拼成 4x4 联系表，便于逐页目视检查"""
import os, win32com.client as win32, pymupdf

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\pages"
os.makedirs(OUT, exist_ok=True)
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "p1"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "p2"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "p3"),
]

w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    for path, tag in DOCS:
        pdf = os.path.join(OUT, tag + ".pdf")
        if not os.path.exists(pdf):
            d = w.Documents.Open(path, ReadOnly=True)
            d.ExportAsFixedFormat(pdf, 17, OpenAfterExport=False, OptimizeFor=0,
                                  Range=0, Item=0, IncludeDocProps=True,
                                  CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True)
            d.Close(False)
        doc = pymupdf.open(pdf)
        n = len(doc)
        print(f"{tag}: {n} 页")
        # 4x4 联系表
        for start in range(0, n, 16):
            pages = doc[start:start + 16]
            cols = 4
            rows = (len(pages) + cols - 1) // cols
            W_, H_ = 420, 594
            sheet = pymupdf.open()
            big = sheet.new_page(width=W_ * cols, height=H_ * rows)
            for i, pg in enumerate(pages):
                r, c = divmod(i, cols)
                rect = pymupdf.Rect(c * W_, r * H_, (c + 1) * W_, (r + 1) * H_)
                big.show_pdf_page(rect, doc, pg.number)
            f = os.path.join(OUT, f"{tag}_sheet_{start//16+1}.png")
            big.get_pixmap(dpi=55).save(f)
            print("  ", f)
        doc.close()
finally:
    w.Quit()
