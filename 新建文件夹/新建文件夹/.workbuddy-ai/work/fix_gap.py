# -*- coding: utf-8 -*-
"""按模板明文修正摘要区"空N行"：摘要前1、摘要→关键词2、关键词→Abstract1、Abstract→Key words2"""
import sys, io, os, copy
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
JOBS = [('摘 要：', 1), ('关键词：', 2), ('Abstract', 1), ('Key words', 2)]

def locate(ps, key):
    for i, p in enumerate(ps):
        if p.text.strip().startswith(key):
            return i
    return None

def clean_empty(src_el):
    el = copy.deepcopy(src_el)
    for a in ('{http://schemas.microsoft.com/office/word/2010/wordml}paraId',
              '{http://schemas.microsoft.com/office/word/2010/wordml}textId'):
        if el.get(a) is not None: del el.attrib[a]
    for br in el.iter(qn('w:br')):
        br.getparent().remove(br)
    for t in el.iter(qn('w:t')):
        t.text = ''
    return el

for nm in FILES:
    path = os.path.join(W, nm, nm + '.docx')
    d = Document(path)
    print('=' * 74); print(nm[:28])
    for key, need in JOBS:
        ps = d.paragraphs
        i = locate(ps, key)
        if i is None:
            print(f'   ?? 未找到 {key}'); continue
        n = 0; j = i - 1
        while j >= 0 and not ps[j].text.strip():
            n += 1; j -= 1
        tgt = ps[i]._element
        if n < need:
            src = ps[i - 1]._element if n > 0 else tgt
            for _ in range(need - n):
                tgt.addprevious(clean_empty(src))
            print(f'   {key:<12} 空行 {n} -> {need}（补 {need-n}）')
        elif n > need:
            for _ in range(n - need):
                tgt.getprevious().getparent().remove(tgt.getprevious())
            print(f'   {key:<12} 空行 {n} -> {need}（删 {n-need}）')
        else:
            print(f'   {key:<12} 空行 {n}（已达标）')
    d.save(path)
print('完成')
