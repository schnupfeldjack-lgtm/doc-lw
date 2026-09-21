import os, pymupdf
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/pages"
for a in [32, 33]:
    doc = pymupdf.open(os.path.join(OUT, "p1.pdf"))
    f = os.path.join(OUT, f"z2_p1_p{a}.png")
    doc[a-1].get_pixmap(dpi=100).save(f); print(f)
    doc.close()
