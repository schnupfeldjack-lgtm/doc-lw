# -*- coding: utf-8 -*-
"""渲染 PDF 为逐页 PNG + 拼版图，供视觉核验。"""
import sys, os
import pymupdf
from PIL import Image

def main(pdf, outdir, zoom=1.4, per_sheet=6):
    os.makedirs(outdir, exist_ok=True)
    doc = pymupdf.open(pdf)
    n = doc.page_count
    paths = []
    for i in range(n):
        pix = doc[i].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
        p = os.path.join(outdir, f'p{i+1:03d}.png')
        pix.save(p)
        paths.append(p)
    # 拼版
    sheets = []
    for start in range(0, n, per_sheet):
        chunk = paths[start:start + per_sheet]
        ims = [Image.open(x) for x in chunk]
        w = max(im.width for im in ims)
        h = max(im.height for im in ims)
        cols = 3
        rows = (len(ims) + cols - 1) // cols
        sheet = Image.new('RGB', (w * cols, h * rows), 'white')
        for k, im in enumerate(ims):
            sheet.paste(im, ((k % cols) * w, (k // cols) * h))
        sp = os.path.join(outdir, f'sheet_{start+1:03d}-{start+len(chunk):03d}.png')
        sheet.save(sp)
        sheets.append(sp)
    print('pages', n)
    for s in sheets:
        print(s)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
