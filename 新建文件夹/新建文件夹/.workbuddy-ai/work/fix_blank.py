# -*- coding: utf-8 -*-
"""删除最后一条参考文献与「致  谢」标题之间的空段落（消除多余空白页）"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]

for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    ps = d.paragraphs
    ack = None
    for i, q in enumerate(ps):
        if q.text.strip() == '致  谢':
            ack = i
            break
    if ack is None:
        print(f'{nm[:18]} 未找到致谢')
        continue
    # 往前找最后一条文献
    last_ref = None
    for j in range(ack - 1, -1, -1):
        if re.match(r'^\[\d+\]', ps[j].text.strip()):
            last_ref = j
            break
    if last_ref is None:
        print(f'{nm[:18]} 未找到文献')
        continue
    # 删除 last_ref+1 .. ack-1 中的空段落
    removed = 0
    for j in range(ack - 1, last_ref, -1):
        if not ps[j].text.strip():
            ps[j]._element.getparent().remove(ps[j]._element)
            removed += 1
    print(f'{nm[:18]} 致谢段={ack} 末条文献={last_ref} 删除空段={removed}')
    d.save(p)
