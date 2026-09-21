# -*- coding: utf-8 -*-
import docx, os, re
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
sub = "再生混凝土在装配式建筑中的应用评价——以住宅项目为例"
d = docx.Document(os.path.join(BASE, sub, sub+".docx"))
for p in d.paragraphs:
    if '\t' in p.text and re.sub(r'[\s.]','',p.text.split('\t')[0].strip()) in ('结论','致谢','参考文献'):
        print("段文本:", repr(p.text))
        for k,r in enumerate(p.runs): print(f"   r{k} {r.text!r}")
        print("-"*50)
