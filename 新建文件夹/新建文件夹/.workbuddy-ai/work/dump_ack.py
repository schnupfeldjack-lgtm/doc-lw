# -*- coding: utf-8 -*-
import docx, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
subs = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
        '装配式施工质量管理问题及优化研究——以市政项目为例',
        'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']
for sub in subs:
    d = docx.Document(os.path.join(BASE, sub, sub+".docx")); ps=d.paragraphs
    hit=False
    for p in ps:
        if p.text.strip()=='致  谢': hit=True; continue
        if hit and p.text.strip():
            print("="*60); print(sub[:22], "字数", len(p.text))
            print(p.text)
            break
