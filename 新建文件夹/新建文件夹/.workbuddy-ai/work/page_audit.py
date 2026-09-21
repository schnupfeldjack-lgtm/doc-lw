# -*- coding: utf-8 -*-
"""逐页版面自检：找空页、孤行标题、缺页眉、断头章节"""
import os, re, pymupdf
import win32com.client as win32

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\pages"
os.makedirs(OUT, exist_ok=True)
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "p1"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "p2"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "p3"),
]
HDR = "炎黄职业技术学院毕业论文"

# 1. 重新导出最新 PDF
w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    for path, tag in DOCS:
        pdf = os.path.join(OUT, tag + ".pdf")
        d = w.Documents.Open(path, ReadOnly=True)
        d.ExportAsFixedFormat(pdf, 17, OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
                              IncludeDocProps=True, CreateBookmarks=0, DocStructureTags=True, BitmapMissingFonts=True)
        d.Close(False)
finally:
    w.Quit()

for path, tag in DOCS:
    pdf = os.path.join(OUT, tag + ".pdf")
    doc = pymupdf.open(pdf)
    n = len(doc)
    print(f"\n{'='*70}\n{tag}  共 {n} 页\n{'='*70}")
    for i in range(n):
        pg = doc[i]
        txt = pg.get_text("text").strip()
        cn = len(re.findall(r'[\u4e00-\u9fff]', txt))
        has_hdr = HDR in txt
        lines = [l.strip() for l in txt.split('\n') if l.strip()]
        first = lines[0][:26] if lines else '(空)'
        last = lines[-1][:26] if lines else '(空)'
        flags = []
        # 正文页（第9页起，除结论/文献/致谢尾页）应有页眉
        if i + 1 >= 9 and not has_hdr and cn > 0:
            flags.append('★缺页眉')
        if cn == 0:
            flags.append('★空页')
        elif cn < 60:
            flags.append('△内容很少')
        # 页面最后一行是标题（孤行）
        if re.match(r'^\d+\s{0,2}[．.]?\s*\d*\s*\S', last) and cn < 400:
            pass
        print(f"  p{i+1:>3} 汉字={cn:>4} 页眉={'有' if has_hdr else '无':<2} | 首行={first:<28} | 末行={last:<28} {' '.join(flags)}")
    doc.close()
