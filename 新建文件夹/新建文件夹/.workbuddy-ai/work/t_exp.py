import os, win32com.client as win32
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
path = BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work"
w = win32.DispatchEx("Word.Application"); w.Visible=False; w.DisplayAlerts=0
d = w.Documents.Open(path, ReadOnly=True)
for a,b in [(5,7),(31,33),(8,9)]:
    out = os.path.join(OUT, f"p1_{a}-{b}.pdf")
    try:
        d.ExportAsFixedFormat(out, 17, Range=3, From=a, To=b, OpenAfterExport=False,
                              OptimizeFor=0, Item=0, IncludeDocProps=True,
                              CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True)
        print("OK", out)
    except Exception as e:
        print("FAIL", a, b, str(e)[:80])
d.Close(False); w.Quit()
