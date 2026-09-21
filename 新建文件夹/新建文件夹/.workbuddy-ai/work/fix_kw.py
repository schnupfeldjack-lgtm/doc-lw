# -*- coding: utf-8 -*-
"""
关键词分隔符对齐模板明文：
  模板：关键词：（小四号宋体、加粗）  ×××，×××，×××，×××（小四号宋体）      -> 中文逗号「，」
  模板：Key words：...×××, ×××, ×××, ×××（小四号Times New Roman）   -> 英文逗号+空格「, 」
"""
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
    n = 0
    for para in d.paragraphs:
        t = para.text.strip()
        if t.startswith('关键词：'):
            full = para.text
            if '；' in full:
                new = full.replace('；', '，')
                if para.runs:
                    para.runs[0].text = new
                    for r in para.runs[1:]:
                        r.text = ''
                n += 1
                print(f'  {nm[:12]} 中文关键词 -> {new.strip()[:56]}')
        if t.startswith('Key words'):
            full = para.text
            if ';' in full:
                new = re.sub(r';\s*', ', ', full)
                new = re.sub(r'(Key words：)\s+', r'\1', new)
                if para.runs:
                    para.runs[0].text = new
                    for r in para.runs[1:]:
                        r.text = ''
                n += 1
                print(f'  {nm[:12]} 英文关键词 -> {new.strip()[:62]}')
    print(f'{nm[:18]} 修正 {n} 处')
    d.save(p)
