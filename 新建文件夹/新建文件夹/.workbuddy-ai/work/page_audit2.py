# -*- coding: utf-8 -*-
"""只报问题的逐页自检（页眉用"第 N 页"判定，避免标题误报）"""
import os, re, pymupdf

OUT = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\pages"
HDR = re.compile(r'炎黄职业技术学院毕业论文\s*第\s*\d+\s*页')
HEAD_LINE = re.compile(r'^(\d+\s{1,2}\S|\d+．\d+\s|\d+\.\d+\.\d+\s|结\s*论$|参\s*考\s*文\s*献$|致\s*谢$)')

for tag in ["p1", "p2", "p3"]:
    pdf = os.path.join(OUT, tag + ".pdf")
    doc = pymupdf.open(pdf)
    n = len(doc)
    print(f"\n===== {tag} ({n} 页) — 仅列出问题 =====")
    for i in range(n):
        pg = doc[i]
        txt = pg.get_text("text").strip()
        cn = len(re.findall(r'[\u4e00-\u9fff]', txt))
        en = len(re.findall(r'[A-Za-z]+', txt))
        has_hdr = bool(HDR.search(txt))
        lines = [l.strip() for l in txt.split('\n') if l.strip()]
        probs = []
        if cn + en == 0:
            probs.append('★空白页')
        # 正文区（第9页起）应有页眉
        if i + 1 >= 9 and not has_hdr and (cn + en) > 0:
            probs.append('★缺页眉')
        # 末行是标题 -> 孤行标题
        if lines:
            last = lines[-1]
            if HEAD_LINE.match(last) and not last.startswith('炎黄'):
                probs.append('★孤行标题:' + last[:16])
        # 内容过少的正文页
        if i + 1 >= 9 and cn < 150 and cn > 0:
            probs.append(f'△内容很少({cn}字)')
        if probs:
            print(f"  p{i+1:>3} 汉字={cn:>4} | {' / '.join(probs)}")
    if not any(True for _ in range(0)):
        pass
    doc.close()
print("\n检查完毕")
