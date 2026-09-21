# -*- coding: utf-8 -*-
"""
模板 vs 论文 逐字段对照。
模板值取「母段落实际值」；同时标注模板「明文规定」，二者冲突时单独列出待确认。
"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
TW = 12700.0
NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TPL = f'{W}\\炎黄职业技术学院毕业论文模板(1).docx'


def ea_of(r):
    try:
        rpr = r._element.rPr
        if rpr is not None and rpr.rFonts is not None:
            return rpr.rFonts.get(NS + 'eastAsia') or rpr.rFonts.get(NS + 'ascii')
    except Exception:
        pass
    return None


def f(para, run_idx=0):
    r = para.runs[run_idx] if len(para.runs) > run_idx else (para.runs[0] if para.runs else None)
    pf = para.paragraph_format
    sz = r.font.size.pt if (r and r.font.size) else None
    ls = pf.line_spacing
    if isinstance(ls, (int, float)) and ls and ls > 10000:
        ls = round(ls / TW, 1)
    sb = pf.space_before
    sa = pf.space_after
    pPr = para._element.find(qn('w:pPr'))
    bl = al_ = None
    if pPr is not None:
        s = pPr.find(qn('w:spacing'))
        if s is not None:
            bl = s.get(qn('w:beforeLines'))
            al_ = s.get(qn('w:afterLines'))
    return dict(sz=sz, ea=ea_of(r) if r else None, b=(r.font.bold if r else None),
                al=para.alignment, ls=ls,
                li=round(pf.left_indent / TW, 2) if pf.left_indent else None,
                fi=round(pf.first_line_indent / TW, 2) if pf.first_line_indent else None,
                sb=round(sb / TW, 1) if sb else None,
                sa=round(sa / TW, 1) if sa else None,
                bl=bl, al_=al_)


from docx.oxml.ns import qn

def show(tag, v):
    return (f"{tag}: sz={v['sz']} 字体={v['ea']} 加粗={v['b']} 对齐={v['al']} "
            f"行距={v['ls']} 左缩进={v['li']} 首行缩进={v['fi']} "
            f"段前={v['sb']}({v['bl']}) 段后={v['sa']}({v['al_']})")


t = Document(TPL)
tp = t.paragraphs

# 模板关键段落索引（来自 dump）
IDX = {
    '摘要': 31, '关键词': 35, 'Abstract': 39, 'Key words': 41,
    '目录标题': 52, '目录一级': 55, '目录二级': 59, '目录无编号': 69,
    '一级标题(章1)': 75, '正文(章首段)': 76, '二级标题': 77, '正文': 78,
    '三级标题': 79, '正文2': 80, '一级标题(章2)': 81, '正文3': 82,
    '结论标题': 90, '结论正文': 92, '参考文献标题': 109, '文献条目': 111,
    '致谢标题': 134, '致谢正文': 136,
}
print('########## 模板母段落实际值 ##########')
for k, i in IDX.items():
    print(f'  {k:<14} {show("", f(tp[i]))}')
    print(f'                文本: {tp[i].text[:50]!r}')

# 论文
FILES = [
    ('论文一', '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'),
    ('论文二', '装配式施工质量管理问题及优化研究——以市政项目为例'),
    ('论文三', 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'),
]
PICK = {
    '摘要': lambda p: p.text.strip().startswith('摘 要：'),
    '关键词': lambda p: p.text.strip().startswith('关键词：'),
    'Abstract': lambda p: p.text.strip().startswith('Abstract'),
    'Key words': lambda p: p.text.strip().startswith('Key words'),
    '目录标题': lambda p: p.text.strip().startswith('目') and '录' in p.text[:4] and len(p.text) < 8,
    '目录一级': lambda p: '\t' in p.text and re.match(r'^\d+\.', p.text.strip()),
    '目录二级': lambda p: '\t' in p.text and re.match(r'^\d+\.\d+', p.text.strip()),
    '目录无编号': lambda p: '\t' in p.text and p.text.strip().split('\t')[0] in ('结论', '致谢', '参考文献'),
    '一级标题(章1)': lambda p: '\t' not in p.text and re.match(r'^1\s+\S', p.text.strip()) and len(p.text) < 40,
    '二级标题': lambda p: '\t' not in p.text and re.match(r'^\d+[.．]\d+\s+\S', p.text.strip()) and len(p.text) < 40,
    '三级标题': lambda p: '\t' not in p.text and re.match(r'^\d+[.．]\d+[.．]\d+\s+\S', p.text.strip()) and len(p.text) < 40,
    '正文': lambda p: '\t' not in p.text and len(p.text) > 120
             and not p.text.strip().startswith(('摘 要', '关键词', 'Abstract', 'Key words', '本文以', '本论文是在'))
             and not re.match(r'^\[\d+\]', p.text.strip()),
    '结论标题': lambda p: p.text.strip() == '结  论',
    '结论正文': lambda p: '\t' not in p.text and p.text.strip().startswith('本文以'),
    '参考文献标题': lambda p: p.text.strip() == '参 考 文 献',
    '文献条目': lambda p: re.match(r'^\[\d+\]', p.text.strip()),
    '致谢标题': lambda p: p.text.strip() == '致  谢',
    '致谢正文': lambda p: '\t' not in p.text and p.text.strip().startswith('本论文是在'),
}

print()
print('########## 论文实测值 ##########')
for tag, nm in FILES:
    d = Document(f'{W}\\{nm}\\{nm}.docx')
    ps = d.paragraphs
    print(f'\n===== {tag} =====')
    for k, pred in PICK.items():
        hits = [p for p in ps if pred(p)]
        if not hits:
            print(f'  {k:<14} 未找到')
            continue
        v = f(hits[0])
        print(f'  {k:<14} {show("", v)}')
