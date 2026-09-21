import fitz, os
OUT = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work"
for name in ["p1_5-7", "p1_8-9", "p1_31-33"]:
    p = os.path.join(OUT, name + ".pdf")
    d = fitz.open(p)
    for i, pg in enumerate(d):
        pix = pg.get_pixmap(dpi=90)
        f = os.path.join(OUT, f"{name}_p{i+1}.png")
        pix.save(f)
        print(f)
    d.close()
