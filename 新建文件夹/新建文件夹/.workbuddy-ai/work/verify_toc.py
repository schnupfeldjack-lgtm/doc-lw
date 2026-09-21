# -*- coding: utf-8 -*-
"""最终校验：Word 真实页码 vs 目录所列页码"""
import re
import win32com.client as win32
from docx import Document

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]

def norm(s):
    return re.sub(r'\s+', '', s or '')

w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    for path, nm in DOCS:
        d = w.Documents.Open(path, ReadOnly=True)
        d.Repaginate()
        total = d.ComputeStatistics(2)
        pmap = {}
        for i in range(1, d.Paragraphs.Count + 1):
            par = d.Paragraphs(i)
            t = norm(par.Range.Text)
            if t and t not in pmap:
                try:
                    pmap[t] = par.Range.Information(3)
                except Exception:
                    pass
        d.Close(False)
        dd = Document(path)
        bad = 0
        rows = []
        for p in dd.paragraphs:
            if '\t' not in p.text:
                continue
            title, _, page = p.text.rpartition('\t')
            if not page.strip().isdigit():
                continue
            real = pmap.get(norm(title))
            flag = 'OK' if real == int(page.strip()) else '★'
            if flag == '★':
                bad += 1
            rows.append((title[:22], page.strip(), real, flag))
        print(f"\n===== {nm}  总页数={total} =====")
        for t, pg, real, fl in rows:
            print(f"  {fl} {t:<24} 目录={pg:<4} 实际={real}")
        print(f"  >>> 不符 {bad} 条")
finally:
    w.Quit()
