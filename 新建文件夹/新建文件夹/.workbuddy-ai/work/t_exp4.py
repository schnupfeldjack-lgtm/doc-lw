import os, win32com.client as win32, pymupdf
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "p1"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "p2"),
]
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/pages"
w = win32.DispatchEx("Word.Application"); w.Visible=False; w.DisplayAlerts=0
try:
    for path, tag in DOCS:
        d = w.Documents.Open(path, ReadOnly=True)
        for a in [8,9]:
            out = os.path.join(OUT, f"v_{tag}_p{a}.pdf")
            d.ExportAsFixedFormat(out, 17, Range=3, From=a, To=a, OpenAfterExport=False, OptimizeFor=0, Item=0)
            doc = pymupdf.open(out)
            for i, pg in enumerate(doc):
                f = os.path.join(OUT, f"v_{tag}_p{a}.png")
                pg.get_pixmap(dpi=110).save(f); print(f)
            doc.close()
        d.Close(False)
finally:
    w.Quit()
