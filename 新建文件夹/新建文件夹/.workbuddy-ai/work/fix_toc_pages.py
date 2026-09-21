# -*- coding: utf-8 -*-
"""用 Word 真实分页重算目录页码（原估算把每段算成 1 页，页码严重虚高）"""
import re, sys, time
import win32com.client as win32
from docx import Document
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]

def norm(s):
    return re.sub(r'\s+', '', s or '')

word = win32.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
try:
    for path, nm in DOCS:
        doc = word.Documents.Open(path, ReadOnly=True)
        doc.Repaginate()
        total = doc.ComputeStatistics(2)   # wdStatisticPages
        pmap = {}
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            t = norm(para.Range.Text)
            if not t:
                continue
            if t not in pmap:
                try:
                    pmap[t] = para.Range.Information(3)  # wdActiveEndPageNumber
                except Exception:
                    pass
        doc.Close(False)
        print(f"\n=== {nm} 总页数={total} 取到页码条目={len(pmap)}")
        d = Document(path)
        changed = miss = 0
        for p in d.paragraphs:
            txt = p.text
            if '\t' not in txt:
                continue
            title, _, page = txt.rpartition('\t')
            if not re.match(r'^[\d．.\s　]*$', '') and not title.strip():
                continue
            if not page.strip().isdigit():
                continue
            key = norm(title)
            real = pmap.get(key)
            if real is None:
                print(f"   [未匹配] {title[:30]!r}")
                miss += 1
                continue
            if int(page.strip()) != real:
                runs = p.runs
                if runs:
                    ttag = runs[-1]._element.find(q('t'))
                    if ttag is not None:
                        ttag.text = str(real)
                        ttag.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        changed += 1
        d.save(path)
        print(f"   目录页码修正 {changed} 条，未匹配 {miss} 条")
finally:
    word.Quit()
print("完成")
