# -*- coding: utf-8 -*-
"""
修「孤行标题 / 标题在页尾而正文在下一页」：
- 所有标题（章/二级/三级/结论/参考文献/致谢）加 keepNext + keepLines（与下段同页、段中不分页）
- 含图片的段落加 keepNext（图与图题不分家）
- 图题/表题加 keepLines
"""
import re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三"),
]

# pPr 子元素顺序：keepNext, keepLines 在 pStyle 之后、pageBreakBefore 之前
def add_flag(pPr, tag):
    if pPr.find(q(tag)) is not None:
        return 0
    e = OxmlElement('w:' + tag)
    # 放在 pStyle 之后（若有），否则放最前
    ps = pPr.find(q('pStyle'))
    if ps is not None:
        ps.addnext(e)
    else:
        pPr.insert(0, e)
    return 1

def get_pPr(p):
    pPr = p._p.find(q('pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p._p.insert(0, pPr)
    return pPr

HEAD_RE = [
    (r'^\d+\s{1,2}\S', '章标题'),                 # 1  绪论
    (r'^\d+．\d+', '二级标题'),                    # 1．1
    (r'^\d+\.\d+\.\d+', '三级标题'),
    (r'^结\s*论$', '结论标题'),
    (r'^参\s*考\s*文\s*献$', '文献标题'),
    (r'^致\s*谢$', '致谢标题'),
]

for path, nm in DOCS:
    d = Document(path)
    n_head = n_fig = 0
    for p in d.paragraphs:
        t = p.text.strip()
        if '\t' in t:          # 目录条目跳过
            continue
        for pat, _lab in HEAD_RE:
            if re.match(pat, t) and len(t) < 60:
                pPr = get_pPr(p)
                n_head += add_flag(pPr, 'keepNext')
                add_flag(pPr, 'keepLines')
                break
        # 图题 / 表题
        if re.match(r'^(图|表)\s*\d+[–\-]\d+\s{2}\S', t) and len(t) < 50:
            add_flag(get_pPr(p), 'keepLines')
        # 含图片的段落
        if p._p.findall(f".//{q('drawing')}"):
            n_fig += add_flag(get_pPr(p), 'keepNext')
    d.save(path)
    print(f"{nm}: 标题加 keepNext/keepLines={n_head}  图片段加 keepNext={n_fig}")
print("完成")
