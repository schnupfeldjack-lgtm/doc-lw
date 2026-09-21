# -*- coding: utf-8 -*-
import docx, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
for sub in ['装配式施工质量管理问题及优化研究——以市政项目为例',
            'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']:
    d = docx.Document(os.path.join(BASE, sub, sub+".docx"))
    print("="*60); print(sub[:24])
    for p in d.paragraphs:
        if '\t' in p.text: print("  ", repr(p.text), "|runs:", [r.text for r in p.runs])
