# -*- coding: utf-8 -*-
"""从 PDF 实测每页文本的实际左右边界（cm），对比模板与论文"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import fitz
OUT = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out'

def measure(pdf, tag, maxpage=None):
    d = fitz.open(pdf)
    print(f"\n===== {tag}  共{len(d)}页 =====")
    print("  页  左边界cm  右边界cm   有效宽cm")
    for i in range(len(d)):
        pg = d[i]
        W = pg.rect.width
        xs0, xs1 = [], []
        for b in pg.get_text('blocks'):
            x0, y0, x1, y1 = b[:4]
            txt = b[4].strip()
            if not txt: continue
            if y0 < 40 or y1 > pg.rect.height - 40:   # 排除页眉页脚
                continue
            xs0.append(x0); xs1.append(x1)
        if not xs0:
            print(f"  {i+1:>3}  (无正文)")
            continue
        l = min(xs0)/72*2.54
        r = (W - max(xs1))/72*2.54
        print(f"  {i+1:>3}   {l:6.2f}    {r:6.2f}    {W/72*2.54 - l - r:6.2f}")
        if maxpage and i+1 >= maxpage: break

measure(os.path.join(OUT,'TPL.pdf'), "模板")
measure(os.path.join(OUT,'P1.pdf'), "论文一", 12)
measure(os.path.join(OUT,'P2.pdf'), "论文二", 12)
measure(os.path.join(OUT,'P3.pdf'), "论文三", 12)
