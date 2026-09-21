# -*- coding: utf-8 -*-
"""检查模板新发现的规定：
1) 注1：图与表应设置在文章中首次提到处附近；图与表应有相应的名称
2) 注3：字体颜色统一黑色
3) 多文献角注写法 [4，5] / [6~8]
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
for nm in FILES:
    d = Document(os.path.join(W, nm, nm + '.docx'))
    ps = d.paragraphs
    print('=' * 76); print(nm[:30])
    # 1) 多文献角注格式
    bad_multi = []
    for i, p in enumerate(ps):
        t = p.text
        if re.match(r'^\[\d+\]', t.strip()):
            continue
        # 找形如 [1][2] 或 [1,2] 或 [1，2]
        for m in re.finditer(r'\[(\d+)\]\[(\d+)\]', t):
            bad_multi.append((i, '连续角注未合并', m.group(0)))
        for m in re.finditer(r'\[(\d+)\s*[,，]\s*(\d+)', t):
            bad_multi.append((i, '多文献角注', m.group(0)))
        for m in re.finditer(r'\[(\d+)\s*[~～-]\s*(\d+)\]', t):
            bad_multi.append((i, '范围角注', m.group(0)))
    print(f'  多文献角注: {len(bad_multi)} 处  {bad_multi[:4]}')
    # 2) 图表题是否有名称
    caps = []
    for i, p in enumerate(ps):
        t = p.text.strip()
        m = re.match(r'^(图|表)\s*\d+[–\-—]\d+\s*(.*)$', t)
        if m and len(t) < 45:
            caps.append((i, m.group(1) + m.group(2)[:0], t[:40]))
    print(f'  图表题: {len(caps)} 个')
    for c in caps:
        print(f'     段{c[0]}: {c[2]}')
    # 3) 图表引用句 vs 图表题位置（首次提到处附近）
    for i, p in enumerate(ps):
        for m in re.finditer(r'(图|表)\s*(\d+)[–\-—](\d+)', p.text):
            pass
    # 找"图X–Y 给出了/所示"的引用段，和对应图题段
    refs = {}
    for i, p in enumerate(ps):
        for m in re.finditer(r'(图|表)\s*(\d+)[–\-—](\d+)', p.text):
            refs.setdefault((m.group(1), m.group(2), m.group(3)), []).append(i)
    for k, v in refs.items():
        kind, ch, no = k
        cap_idx = [i for i, p in enumerate(ps)
                   if re.match(rf'^{kind}\s*{ch}[–\-—]{no}\s', p.text.strip())]
        print(f'   {kind}{ch}–{no}: 提及于段{v}  题注于段{cap_idx}')
