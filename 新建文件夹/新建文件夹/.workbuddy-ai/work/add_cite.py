# -*- coding: utf-8 -*-
"""在正文补角注（上标），使每条文献都有对应引用；论文三同时做文献表重排与重编号"""
import sys, io, os, re, copy
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
P2 = '装配式施工质量管理问题及优化研究——以市政项目为例'
P3 = 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'

def set_text(run_el, txt):
    for t in run_el.findall(qn('w:t')):
        run_el.remove(t)
    t = OxmlElement('w:t')
    t.text = txt
    t.set(qn('xml:space'), 'preserve')
    run_el.append(t)

def mk_run(p, txt, sup=False):
    src = p.runs[-1]._element
    el = copy.deepcopy(src)
    set_text(el, txt)
    if sup:
        rPr = el.find(qn('w:rPr'))
        if rPr is None:
            rPr = OxmlElement('w:rPr'); el.insert(0, rPr)
        for v in rPr.findall(qn('w:vertAlign')):
            rPr.remove(v)
        va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rPr.append(va)
    return el

def append_cite(p, words, num):
    p._element.append(mk_run(p, words, sup=False))
    p._element.append(mk_run(p, f'[{num}]', sup=True))
    print(f'    + 段尾追加: {words[:28]}… [{num}]')

def insert_after(p, marker, words, num):
    for r in p.runs:
        if marker in r.text:
            r._element.addnext(mk_run(p, words, sup=False))
            r._element.addnext(mk_run(p, f'[{num}]', sup=True))
            print(f'    + 段中插入于 {marker!r} 后: {words[:24]}… [{num}]')
            return True
    return False

# ---------- 论文二 ----------
path2 = os.path.join(W, P2, P2 + '.docx')
d = Document(path2); ps = d.paragraphs
print('=' * 72); print('论文二：补 3 处角注')
append_cite(ps[100], '《装配式建筑评价标准》对装配率的计算方法和评价等级作出了规定', 10)
append_cite(ps[119], '《建筑信息模型应用统一标准》则为BIM在工程中的统一应用框架提供了依据', 11)
append_cite(ps[191], '相关研究表明，BIM在施工项目管理中的应用能够显著提升质量管理的精细化水平', 12)
d.save(path2)

# ---------- 论文三 ----------
path3 = os.path.join(W, P3, P3 + '.docx')
d = Document(path3); ps = d.paragraphs
print('=' * 72); print('论文三：补 2 处角注 + 文献表重排')
append_cite(ps[97], '已有研究对装配式建筑质量影响因素的识别和控制方法作了系统分析', 11)
ok = insert_after(ps[103], '[3]', '预制装配式建筑施工技术的研究成果为施工环节的技术应用提供了参考', 12)
print('   段103 插入成功:', ok)

# 角注重编号 8->7, 9->8, 10->9, 11->10
MAPPING = {8: 7, 9: 8, 10: 9, 11: 10}
cnt = 0
for p in d.paragraphs:
    if re.match(r'^\[\d+\]', p.text.strip()):
        continue
    for r in p.runs:
        if '[' in r.text:
            new = re.sub(r'\[(\d+)\]',
                         lambda m: f'[{MAPPING.get(int(m.group(1)), int(m.group(1)))}]', r.text)
            if new != r.text:
                r.text = new; cnt += 1
print(f'   正文角注重编号 run 数: {cnt}')
# 文献表重排：原[7]杨婷 -> 位置11, 原[12]崔宇 -> 位置12
refs = [p for p in d.paragraphs if re.match(r'^\[\d+\]', p.text.strip())]
old_texts = []
for p in refs:
    n = int(re.match(r'^\[(\d+)\]', p.text.strip()).group(1))
    old_texts.append((n, p))
byidx = {n: p for n, p in old_texts}
neworder = [1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 7, 12, 13, 14, 15]
assert sorted(neworder) == sorted(byidx.keys())
for new, old in enumerate(neworder, 1):
    p = byidx[old]
    for r in p.runs:
        if re.match(r'^\[\d+\]', r.text):
            r.text = re.sub(r'^\[\d+\]', f'[{new}]', r.text)
            break
print('   文献表新顺序:', neworder)
d.save(path3)
print('完成')
