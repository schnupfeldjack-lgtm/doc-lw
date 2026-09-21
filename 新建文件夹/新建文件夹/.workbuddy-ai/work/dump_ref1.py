# -*- coding: utf-8 -*-
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
nm = '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'
d = Document(f'{W}\{nm}\{nm}.docx')
for p in d.paragraphs:
    t = p.text.strip()
    if re.match(r'^\[\d+\]', t):
        print(f'{len(t):3d} | {t}')
