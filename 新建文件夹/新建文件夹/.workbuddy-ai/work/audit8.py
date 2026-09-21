# -*- coding: utf-8 -*-
"""检查 PDF 中图形/表格线是否超出正文版心（左>3.18cm 或 右<3.17cm 视为越界）"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import fitz
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'
PAGE_W_CM = 21.0
for tag in ('P1','P2','P3'):
    d = fitz.open(os.path.join(OUT, tag+'.pdf'))
    bad = []
    for i in range(len(d)):
        pg = d[i]
        if i < 7: continue          # 前面是封面/目录节，版心不同
        for dr in pg.get_drawings():
            x0 = dr['rect'].x0/72*2.54; x1 = dr['rect'].x1/72*2.54
            if x0 < 3.0 or x1 > PAGE_W_CM - 3.0:
                bad.append((i+1, round(x0,2), round(x1,2)))
        for b in pg.get_text('blocks'):
            if b[0]/72*2.54 < 3.0 or b[2]/72*2.54 > PAGE_W_CM - 3.0:
                bad.append((i+1, round(b[0]/72*2.54,2), round(b[2]/72*2.54,2), b[4][:20]))
    print(f"{tag}: 越界 {len(bad)} 处")
    for x in bad[:8]: print("   ", x)
