import os, pymupdf
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/pages"
for tag, pages in [("p1",[30,31,32]), ("p2",[29,30,31]), ("p3",[29,30,31])]:
    doc = pymupdf.open(os.path.join(OUT, tag + ".pdf"))
    for a in pages:
        f = os.path.join(OUT, f"z_{tag}_p{a}.png")
        doc[a-1].get_pixmap(dpi=100).save(f); print(f)
    doc.close()
