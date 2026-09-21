# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import fitz
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'
W = 21.0
for tag in ('P1','P2','P3','TPL'):
    d = fitz.open(os.path.join(OUT, tag+'.pdf'))
    print(f"\n===== {tag} (只看矢量图形/表格线) =====")
    for i in range(len(d)):
        pg = d[i]
        xs0, xs1 = [], []
        for dr in pg.get_drawings():
            r = dr['rect']
            if r.width < 5: continue
            xs0.append(r.x0/72*2.54); xs1.append(r.x1/72*2.54)
        if not xs0: continue
        l, r = min(xs0), W - max(xs1)
        flag = "  <== 超界" if (l < 3.0 or r < 3.0) else ""
        print(f"  p{i+1}: 左{l:6.2f} 右{r:6.2f}{flag}")
