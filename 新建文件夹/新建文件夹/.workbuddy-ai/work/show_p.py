# -*- coding: utf-8 -*-
import sys, io, fitz
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
d = fitz.open(r'C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\out\P1.pdf')
for i in (28, 29, 30, 31, 32):
    print(f'########## P1 第{i+1}页 ##########')
    print(d[i].get_text('text'))
