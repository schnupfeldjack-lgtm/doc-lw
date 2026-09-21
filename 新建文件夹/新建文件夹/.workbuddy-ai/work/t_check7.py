import pymupdf
doc = pymupdf.open(r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/pages/p1.pdf")
pg = doc[6]  # page 7
print(pg.get_text("text"))
