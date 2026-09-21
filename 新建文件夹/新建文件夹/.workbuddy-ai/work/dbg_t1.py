from docx import Document
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
path = BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"
d = Document(path)
for ti in [2, 3]:
    tb = d.tables[ti]
    print(f"\n=== 表{ti} ===")
    seen = set()
    for ri, row in enumerate(tb.rows):
        cells = []
        for c in row.cells:
            if id(c._tc) in seen: continue
            seen.add(id(c._tc))
            cells.append(c.text.replace('\n', ' ').strip())
        print(f"  R{ri}: {cells}")
