# -*- coding: utf-8 -*-
"""把 **xxx** 形式的 Markdown 残留转成真正的 Word 加粗 run"""
import sys, io, re, copy
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]
PAT = re.compile(r'\*\*(.+?)\*\*')

def boldify(para):
    """把段落里的 **x** 转成加粗；返回是否改动"""
    full = para.text
    if '**' not in full:
        return False
    # 收集 (run_index, char_index) -> 逐字符切分
    pieces = []           # (char, run_idx)
    for ri, r in enumerate(para.runs):
        for ch in r.text:
            pieces.append((ch, ri))
    # 在字符序列上做正则
    s = ''.join(c for c, _ in pieces)
    spans = []            # (start, end) 需要加粗的区间
    pos = 0
    out_chars = []
    for m in PAT.finditer(s):
        # m.group(1) 内容 -> 加粗
        pass
    # 重建：去掉 **，记录加粗区间
    new_pieces = []
    i = 0
    while i < len(s):
        if s.startswith('**', i):
            j = s.find('**', i + 2)
            if j == -1:
                new_pieces.append((s[i], pieces[i][1]))
                i += 1
                continue
            inner = s[i + 2:j]
            for k in range(i + 2, j):
                new_pieces.append((s[k], pieces[k][1], True))
            i = j + 2
        else:
            new_pieces.append((s[i], pieces[i][1]))
            i += 1
    # 按 run_idx 分组重排
    runs = para.runs
    if not runs:
        return False
    # 目标：重建 runs。基础 rPr 取原 run 的 rPr
    from docx.oxml import OxmlElement
    new_texts = []        # list of (text, bold, run_idx)
    for item in new_pieces:
        ch = item[0]
        ri = item[1]
        bold = item[2] if len(item) > 2 else False
        if new_texts and new_texts[-1][1] == bold and new_texts[-1][2] == ri:
            new_texts[-1][0] += ch
        else:
            new_texts.append([ch, bold, ri])
    # 保留第一个 run 作为模板，删除其余
    tmpl = runs[0]
    tmpl_rpr = copy.deepcopy(tmpl._element.rPr) if tmpl._element.rPr is not None else None
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    parent = tmpl._element.getparent()
    tmpl._element.getparent().remove(tmpl._element)
    for txt, bold, ri in new_texts:
        el = copy.deepcopy(tmpl._element) if tmpl is not None else None
        # 清掉所有 w:t
        for t in el.findall(qn('w:t')):
            el.remove(t)
        t = OxmlElement('w:t')
        t.text = txt
        t.set(qn('xml:space'), 'preserve')
        # w:t 必须放在 rPr 之后
        rpr = el.find(qn('w:rPr'))
        if rpr is not None:
            rpr.addnext(t)
        else:
            el.insert(0, t)
        if bold:
            rpr = el.find(qn('w:rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                el.insert(0, rpr)
            b = rpr.find(qn('w:b'))
            if b is None:
                b = OxmlElement('w:b')
                rpr.insert(0, b)
            b.set(qn('w:val'), '1')
            # 顺序：b/bCs 应在 rFonts 之后 —— 用规范顺序重排
        else:
            rpr = el.find(qn('w:rPr'))
            if rpr is not None:
                b = rpr.find(qn('w:b'))
                if b is not None:
                    rpr.remove(b)
        parent.append(el)
    return True

for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    n = 0
    for para in d.paragraphs:
        if '**' in para.text:
            if boldify(para):
                n += 1
    print(f'{nm[:20]} 处理 {n} 段')
    d.save(p)

# 复核
print()
for nm in FILES:
    d = Document(f'{W}\\{nm}\\{nm}.docx')
    left = [p.text[:40] for p in d.paragraphs if '**' in p.text]
    print(f'{nm[:16]} 残留 {len(left)} 处 {left[:2]}')
