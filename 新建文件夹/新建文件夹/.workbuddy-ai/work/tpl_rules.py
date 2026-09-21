# -*- coding: utf-8 -*-
"""提取模板中所有明文格式规定段落（按 python-docx 段落索引）"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
d = Document(os.path.join(W, '炎黄职业技术学院毕业论文模板(1).docx'))
PAT = re.compile(r'[（(][^）)]*(\d+\s*号|[0-9.]+\s*倍行距|磅|加粗|居中|黑体|宋体|Times)[^）)]*[）)]')
for i, p in enumerate(d.paragraphs):
    t = p.text.strip()
    if not t: continue
    if PAT.search(t) or re.match(r'^（空\d行）$', t) or '本页为独立页' in t:
        print(f'TPL[{i:3d}] {t[:96]}')
