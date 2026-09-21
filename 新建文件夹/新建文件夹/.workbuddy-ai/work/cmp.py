import os, pymupdf
OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\pages"
pairs = [(1,1,"01开题报告"),(2,2,"02任务书"),(3,3,"03中期检查"),(4,4,"04封面"),(5,5,"05摘要"),(6,6,"06Abstract"),(7,7,"07目录"),(8,9,"08正文首页"),(11,32,"33致谢")]
W_, H_ = 500, 710
for t, g, lab in pairs:
    out = os.path.join(OUT, f"cmp_{lab}.png")
    t_doc = pymupdf.open(os.path.join(OUT, "tpl.pdf"))
    g_doc = pymupdf.open(os.path.join(OUT, "gen.pdf"))
    sheet = pymupdf.open()
    p = sheet.new_page(width=W_ * 2 + 30, height=H_ + 30)
    p.show_pdf_page(pymupdf.Rect(10, 10, 10 + W_, 10 + H_), t_doc, t-1)
    p.show_pdf_page(pymupdf.Rect(20 + W_, 10, 20 + 2 * W_, 10 + H_), g_doc, g-1)
    sheet[0].get_pixmap(dpi=72).save(out)
    sheet.close(); t_doc.close(); g_doc.close()
    print(lab)
