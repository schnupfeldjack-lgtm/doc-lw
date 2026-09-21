# -*- coding: utf-8 -*-
"""模板明文：应用两篇及以上文献论述同一观点时，角注写成 [4，5]（中文逗号）
   把连写的 [15][16] 合并为 [15，16]"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
for nm in FILES:
    path = os.path.join(W, nm, nm + '.docx')
    d = Document(path)
    n = 0
    for p in d.paragraphs:
        if re.match(r'^\[\d+\]', p.text.strip()):
            continue
        runs = p.runs
        for i in range(len(runs) - 1):
            a, b = runs[i].text, runs[i + 1].text
            if re.fullmatch(r'\[\d+\]', a) and re.fullmatch(r'\[\d+\]', b):
                x, y = a[1:-1], b[1:-1]
                runs[i].text = f'[{x}，'
                runs[i + 1].text = f'{y}]'
                n += 1
                print(f'   {nm[:14]} 段: ...{p.text[max(0,p.text.find(x)-18):p.text.find(y)+4]}')
    d.save(path)
    print(f'{nm[:24]}  合并 {n} 处')
