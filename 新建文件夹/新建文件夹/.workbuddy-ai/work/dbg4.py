# -*- coding: utf-8 -*-
import re
from docx import Document
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]
for path, nm in DOCS:
    d = Document(path)
    for p in d.paragraphs:
        t = p.text.strip()
        if re.match(r'^摘\s*要[：:]', t):
            body = t.split('：', 1)[1]
            print(f"\n[{nm}] 汉字={len(re.findall(r'[\u4e00-\u9fff]', body))} 总字符={len(body.replace(' ',''))}")
            print(body)
