# -*- coding: utf-8 -*-
"""清除参考文献/致谢标题前"空2行"段落里残留的分页符（这些空行是我补的，模板无强制分页）"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn

W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
for nm in FILES:
    path = os.path.join(W, nm, nm + '.docx')
    d = Document(path); ps = d.paragraphs
    print('=' * 70); print(nm[:28])
    for key in ('参 考 文 献', '致  谢'):
        idx = [i for i, p in enumerate(ps) if p.text.strip().startswith(key[:2])
               and key.replace(' ', '') in p.text.replace(' ', '')]
        for i in idx:
            for j in (i - 1, i - 2):
                if j < 0 or ps[j].text.strip():
                    continue
                el = ps[j]._element
                n = 0
                for br in el.iter(qn('w:br')):
                    t = br.get(qn('w:type'))
                    if t == 'page':
                        br.getparent().remove(br); n += 1
                if n:
                    print(f'   {key} 前空段[{j}] 删除分页符 {n} 个')
    d.save(path)
print('完成')
