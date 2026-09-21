# -*- coding: utf-8 -*-
"""按模板明文要求逐条对照生成文档，输出通过/不通过报告"""
import os, re
import docx
from docx.oxml.ns import qn

# 模板明文要求：(字号半点, 中文字体, 是否加粗, 对齐, 行距, 行距规则)
REQ = {
    '章标题':     ('30', '黑体', True,  'left',   '360', 'auto'),
    '二级标题':   ('28', '黑体', True,  'left',   '360', 'auto'),
    '三级标题':   ('24', '黑体', False, 'left',   '360', 'auto'),
    '正文':       ('24', '宋体', False, 'left',   '360', 'auto'),
    '摘要':       ('24', '宋体', True,  'left',   '560', 'exact'),
    '关键词':     ('24', '宋体', True,  'left',   '360', 'auto'),
    'Abstract':   ('24', 'Times New Roman', True,  'left', '360', 'auto'),
    'Key words':  ('24', 'Times New Roman', True,  'left', '360', 'auto'),
    '目录标题':   ('36', '黑体', True,  'center', '360', 'auto'),
    '目录条目':   ('24', '宋体', False, 'left',   '560', 'exact'),
    '结论标题':   ('28', '宋体', True,  'center', '360', 'auto'),
    '结论正文':   ('24', '宋体', False, 'left',   '440', 'exact'),
    '文献标题':   ('28', '宋体', True,  'center', '360', 'auto'),
    '文献条目':   ('21', '宋体', False, 'left',   '440', 'exact'),
    '致谢标题':   ('28', '宋体', True,  'left',   '360', 'auto'),
    '致谢正文':   ('24', '宋体', False, 'left',   '440', 'exact'),
    '图表题':     ('21', '宋体', False, 'center', None,  None),
}

# 定位各类段落的谓词（排除目录里的 tab 行）
FIND = {
    '章标题':   lambda t: re.match(r'^\d\s{2}\S', t) is not None,
    '二级标题': lambda t: re.match(r'^\d．\d\s', t) is not None,
    '三级标题': lambda t: re.match(r'^\d+\.\d+\.\d+\s', t) is not None,
    # 正文：长段落，且排除摘要/关键词/Abstract/Key words（它们也是长段落，否则会被误判为正文）
    '正文':     lambda t: (len(t) > 120 and not re.match(r'^\d', t)
                           and not t.startswith('[') and not t.startswith('（')
                           and not t.startswith('摘 要') and not t.startswith('关键词')
                           and not t.startswith('Abstract') and not t.startswith('Key words')),
    '摘要':     lambda t: t.startswith('摘 要：'),
    '关键词':   lambda t: t.startswith('关键词：'),
    'Abstract': lambda t: t.startswith('Abstract：'),
    'Key words': lambda t: t.startswith('Key words：'),
    '目录标题': lambda t: t.strip() == '目  录',
    '目录条目': lambda t: bool(re.match(r'^\d', t)) and '\t' in t,
    '结论标题': lambda t: t.strip() == '结  论',
    '结论正文': lambda t: t.startswith('本文以'),
    '文献标题': lambda t: t.strip() == '参 考 文 献',
    '文献条目': lambda t: t.startswith('[') and len(t) > 20,
    '致谢标题': lambda t: t.strip() == '致  谢',
    '致谢正文': lambda t: t.startswith('本论文是在指导老师'),
    '图表题':   lambda t: bool(re.match(r'^(图|表)\d+[-–]\d+  ', t)),
}


def fmt_of(p):
    x = p._p
    ppr = x.find(qn('w:pPr'))
    px = ppr.xml if ppr is not None else ''
    rx = ''
    for r in x.findall(qn('w:r')):
        if r.find(qn('w:t')) is not None:
            rx = r.xml
            break
    g = lambda pat, s: (re.search(pat, s) or [None, None])[1]
    return {
        'sz': g(r'<w:sz w:val="(\d+)"', rx),
        'ea': g(r'w:eastAsia="([^"]+)"', rx),
        'jc': g(r'<w:jc w:val="([^"]+)"', px) or 'left',
        'b': '<w:b/>' in rx,
        'line': g(r'w:line="(\d+)"', px),
        'rule': g(r'w:lineRule="([^"]+)"', px),
    }


def check(path, label):
    d = docx.Document(path)
    ps = [p for p in d.paragraphs if p.text.strip()]
    print('=' * 74)
    print('格式对照报告：', label)
    print('=' * 74)
    bad = 0
    for key, (sz, ea, b, jc, line, rule) in REQ.items():
        pred = FIND[key]
        hit = None
        for p in ps:
            t = p.text.strip()
            if key != '目录条目' and '\t' in t:
                continue
            if pred(t):
                hit = p
                break
        if hit is None:
            print(f'  [缺失] {key:8s} —— 未找到该类型段落')
            bad += 1
            continue
        f = fmt_of(hit)
        diffs = []
        if f['sz'] != sz:
            diffs.append(f"字号{f['sz']}≠{sz}")
        if ea and f['ea'] != ea:
            diffs.append(f"字体{f['ea']}≠{ea}")
        if f['b'] != b:
            diffs.append(f"加粗{f['b']}≠{b}")
        if f['jc'] != jc:
            diffs.append(f"对齐{f['jc']}≠{jc}")
        if line and f['line'] != line:
            diffs.append(f"行距{f['line']}/{f['rule']}≠{line}/{rule}")
        if diffs:
            print(f'  [不符] {key:8s} {"; ".join(diffs)}')
            bad += 1
        else:
            print(f'  [通过] {key:8s} sz={f["sz"]} {f["ea"]} 粗={f["b"]} {f["jc"]}') 
    # 颜色 & 残留
    alltext = '\n'.join(p.text for p in d.paragraphs)
    blue = sum(1 for p in d.paragraphs for r in p.runs
               if re.search(r'<w:color w:val="(0000FF|FF0000)"', r._r.xml))
    print(f'  [{"通过" if blue==0 else "不符"}] 全文颜色  非黑 run={blue}')
    if blue:
        bad += 1
    zh = len([c for c in alltext if '\u4e00' <= c <= '\u9fff'])
    print(f'  [信息] 全文中文字数 {zh}')
    print(f'  —— 合计不符 {bad} 项')
    print()
    return bad


if __name__ == '__main__':
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    papers = [
        ('再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx', '论文一'),
        ('装配式施工质量管理问题及优化研究——以市政项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例.docx', '论文二'),
        ('BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx', '论文三'),
    ]
    total = 0
    for d, f, lab in papers:
        p = os.path.join(root, d, f)
        total += check(p, lab)
    print('三篇合计不符项：', total)
