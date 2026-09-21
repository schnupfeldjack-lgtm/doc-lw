# -*- coding: utf-8 -*-
"""导出模板与论文 PDF，渲染关键页并拼接并排对照图"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
import pymupdf as fitz

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'
os.makedirs(OUT, exist_ok=True)
TPL = f'{W}\\炎黄职业技术学院毕业论文模板(1).docx'
PAPERS = [
    ('P1', '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'),
    ('P2', '装配式施工质量管理问题及优化研究——以市政项目为例'),
    ('P3', 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'),
]

app = wc.gencache.EnsureDispatch('Word.Application')
app.Visible = False
app.DisplayAlerts = 0


def to_pdf(docx, pdf):
    d = app.Documents.Open(docx, ReadOnly=True)
    d.ExportAsFixedFormat(pdf, 17)
    n = d.ComputeStatistics(2)
    d.Close(False)
    return n


tpl_pdf = f'{OUT}\\TPL.pdf'
nt = to_pdf(TPL, tpl_pdf)
print(f'模板 {nt} 页')
pn = {}
for tag, nm in PAPERS:
    pn[tag] = to_pdf(f'{W}\\{nm}\\{nm}.docx', f'{OUT}\\{tag}.pdf')
    print(f'{tag} {pn[tag]} 页')
app.Quit()


def page_png(pdf, idx, out, zoom=1.6):
    d = fitz.open(pdf)
    pg = d[idx]
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(out)
    r = (pix.width, pix.height)
    d.close()
    return r


def side_by_side(l, r, out, labels):
    dl = fitz.open(l[0]); dr = fitz.open(r[0])
    pl = dl[l[1]].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
    pr = dr[r[1]].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
    Wd = pl.width + pr.width + 40
    Ht = max(pl.height, pr.height) + 46
    doc = fitz.open()
    pg = doc.new_page(width=Wd, height=Ht)
    pg.insert_image(fitz.Rect(0, 34, pl.width, 34 + pl.height), pixmap=pl)
    pg.insert_image(fitz.Rect(pl.width + 40, 34, pl.width + 40 + pr.width,
                              34 + pr.height), pixmap=pr)
    pg.insert_text((10, 22), labels[0], fontsize=13, color=(0.1, 0.3, 0.8))
    pg.insert_text((pl.width + 50, 22), labels[1], fontsize=13, color=(0.8, 0.2, 0.2))
    doc.save(out)
    doc.close(); dl.close(); dr.close()
    print('  ->', os.path.basename(out))


# 对照组：(模板页, 论文页) 0-based
GROUPS = [
    ('摘要', 4, 4),
    ('目录', 6, 6),
    ('正文首页', 8, 8),
]
for name, tp, pp in GROUPS:
    for tag, nm in PAPERS[:1]:
        out = f'{OUT}\\cmp_{name}.png'
        side_by_side((tpl_pdf, tp), (f'{OUT}\\{tag}.pdf', pp), out,
                     (f'模板 {name}页 p{tp+1}', f'论文一 {name}页 p{pp+1}'))
