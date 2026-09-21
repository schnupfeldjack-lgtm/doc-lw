# -*- coding: utf-8 -*-
import docx, os, re
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
for sub, idxs in [("再生混凝土在装配式建筑中的应用评价——以住宅项目为例",[(144,152)]),
                  ("装配式施工质量管理问题及优化研究——以市政项目为例",[(128,134)]),
                  ("BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例",[(167,175)])]:
    d = docx.Document(os.path.join(BASE, sub, sub+".docx")); ps=d.paragraphs
    print("="*70); print(sub[:26])
    for a,b in idxs:
        for j in range(a,b):
            if j < len(ps): print(f"  [{j}] {ps[j].text[:150]}")
