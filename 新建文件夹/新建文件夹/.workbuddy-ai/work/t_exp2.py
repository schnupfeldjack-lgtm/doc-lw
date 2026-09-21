import os, win32com.client as win32, fitz
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
path = BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work"
w = win32.DispatchEx("Word.Application"); w.Visible=False; w.DisplayAlerts=0
d = w.Documents.Open(path, ReadOnly=True)
for a,b in [(5,5),(9,9),(30,32)]:
    out = os.path.join(OUT, f"f_{a}-{b}.pdf")
    d.ExportAsFixedFormat(out, 17, Range=3, From=a, To=b, OpenAfterExport=False,
                          OptimizeFor=0, Item=0, IncludeDocProps=True,
                          CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True)
    doc = fitz.open(out)
    for i, pg in enumerate(doc):
        f = os.path.join(OUT, f"f_{a}-{b}_p{i+1}.png")
        pg.get_pixmap(dpi=88).save(f); print(f)
    doc.close()
d.Close(False); w.Quit()
