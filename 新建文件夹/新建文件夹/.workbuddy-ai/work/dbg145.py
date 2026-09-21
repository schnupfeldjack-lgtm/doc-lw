# -*- coding: utf-8 -*-
import docx, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
sub = "再生混凝土在装配式建筑中的应用评价——以住宅项目为例"
d = docx.Document(os.path.join(BASE, sub, sub+".docx")); ps=d.paragraphs
print("含图4–1?", '图4–1' in ps[145].text)
print("段145 全文:\n", ps[145].text)
print("--- runs ---")
for k,r in enumerate(ps[145].runs): print(k, repr(r.text[-80:]))
