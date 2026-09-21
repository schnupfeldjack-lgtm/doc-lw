# -*- coding: utf-8 -*-
import docx, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
for sub in ["再生混凝土在装配式建筑中的应用评价——以住宅项目为例",
            "BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例"]:
    d = docx.Document(os.path.join(BASE, sub, sub+".docx"))
    ps=d.paragraphs
    idx=[i for i,p in enumerate(ps) if p.text.strip().replace(" ","")=="参考文献"]
    start=idx[-1]
    print("="*70); print(sub, "参考文献标题段:", start)
    for i in range(start+1, min(start+25,len(ps))):
        t=ps[i].text.strip()
        if not t: continue
        if t.startswith("致"): break
        print(f"  [{i}] {t[:160]}")
