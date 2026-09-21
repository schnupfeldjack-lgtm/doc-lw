# -*- coding: utf-8 -*-
"""模板 vs 三篇论文 关键页并排对照图"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pymupdf as fitz
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'
CMP = os.path.join(OUT, 'cmp'); os.makedirs(CMP, exist_ok=True)

def side(l, r, out, labels):
    dl = fitz.open(l[0]); dr = fitz.open(r[0])
    pl = dl[l[1]].get_pixmap(matrix=fitz.Matrix(1.45, 1.45))
    pr = dr[r[1]].get_pixmap(matrix=fitz.Matrix(1.45, 1.45))
    Wd = pl.width + pr.width + 40; Ht = max(pl.height, pr.height) + 46
    doc = fitz.open(); pg = doc.new_page(width=Wd, height=Ht)
    pg.insert_image(fitz.Rect(0, 34, pl.width, 34 + pl.height), pixmap=pl)
    pg.insert_image(fitz.Rect(pl.width + 40, 34, pl.width + 40 + pr.width, 34 + pr.height), pixmap=pr)
    pg.insert_text((10, 22), labels[0], fontsize=12, color=(0.1, 0.3, 0.8))
    pg.insert_text((pl.width + 50, 22), labels[1], fontsize=12, color=(0.8, 0.2, 0.2))
    doc.save(out); doc.close(); dl.close(); dr.close()
    print('  ->', os.path.basename(out))

TPL = os.path.join(OUT, 'TPL.pdf')
# (名称, 模板页0based, P1页, P2页, P3页)
GROUPS = [
    ('1摘要',    4,  4,  4,  4),
    ('2英文摘要', 5,  5,  5,  5),
    ('3目录',    6,  6,  6,  6),
    ('4正文',    7,  8,  8,  8),
    ('5图表',    7, 18, 17, 20),
    ('6结论',    8, 29, 27, 27),
    ('7参考文献', 9, 30, 28, 28),
    ('8致谢',   10, 31, 29, 29),
]
for name, tp, a, b, c in GROUPS:
    for tag, pp in (('P1', a), ('P2', b), ('P3', c)):
        pdf = os.path.join(OUT, tag + '.pdf')
        d = fitz.open(pdf)
        if pp >= len(d):
            d.close(); continue
        d.close()
        side((TPL, tp), (pdf, pp), os.path.join(CMP, f'{name}_{tag}.png'),
             (f'模板·{name[1:]} (p{tp+1})', f'{tag}·{name[1:]} (p{pp+1})'))
print('完成 ->', CMP)
