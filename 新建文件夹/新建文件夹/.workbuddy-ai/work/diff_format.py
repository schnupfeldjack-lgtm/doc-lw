# -*- coding: utf-8 -*-
"""逐字段对比：模板母段落 pPr/rPr  vs  生成文档对应段落 pPr/rPr（忽略颜色）"""
import os, re
import docx
from docx.oxml.ns import qn

TPL = os.path.abspath(os.path.join(
    os.path.dirname(__file__), '..', '..', '炎黄职业技术学院毕业论文模板(1).docx'))


def clean_xml(x):
    """去掉颜色属性（模板是蓝色说明文字，生成文档按模板要求改黑，属预期差异）"""
    x = re.sub(r'<w:color[^>]*/>', '', x)
    return x


def ppr_rpr(p):
    x = p._p
    ppr = x.find(qn('w:pPr'))
    ppr_s = ppr.xml if ppr is not None else '(无pPr)'
    # 取第一个含 w:t 的 run 的 rPr
    rpr_s = '(无run)'
    for r in x.findall(qn('w:r')):
        if r.find(qn('w:t')) is not None:
            rpr = r.find(qn('w:rPr'))
            rpr_s = rpr.xml if rpr is not None else '(无rPr)'
            break
    return clean_xml(ppr_s), clean_xml(rpr_s)


def get(d, pred, exclude_tab=True):
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if exclude_tab and '\t' in t:
            continue
        if pred(t):
            return p
    return None


def load_tpl():
    d = docx.Document(TPL)
    m = {}
    m['H1'] = get(d, lambda t: t.startswith('1  引言'))
    m['H2'] = get(d, lambda t: '正文2级标题' in t)
    m['H3'] = get(d, lambda t: '正文3级标题' in t)
    m['BODY'] = get(d, lambda t: '小4号宋体，1.5倍行距' in t)
    m['ABS'] = get(d, lambda t: t.startswith('摘 要：'))
    m['KW'] = get(d, lambda t: t.startswith('关键词：'))
    m['ABSEN'] = get(d, lambda t: t.startswith('Abstract：'))
    m['KWEN'] = get(d, lambda t: t.startswith('Key words：'))
    m['TOCT'] = get(d, lambda t: '目  录' in t and '黑体小二号' in t)
    m['TOC1'] = get(d, lambda t: t.startswith('1.引言'), exclude_tab=False)
    m['TOC2'] = get(d, lambda t: re.match(r'^4\.1', t), exclude_tab=False)
    m['CONCT'] = get(d, lambda t: t.startswith('结  论') and '四号宋体' in t)
    m['CONCB'] = get(d, lambda t: '小四号宋体，行距为固定值22磅' in t and '结论' not in t)
    m['REFT'] = get(d, lambda t: '参 考 文 献' in t and '四号宋体' in t)
    m['REF'] = get(d, lambda t: t.startswith('[1]'))
    m['ACKT'] = get(d, lambda t: t.startswith('致  谢'))
    m['ACKB'] = None
    # 致谢正文：结论正文之后的下一个同类
    found_first = False
    for p in d.paragraphs:
        t = p.text.strip()
        if '小四号宋体，行距为固定值22磅' in t:
            if not found_first:
                found_first = True
            else:
                m['ACKB'] = p
                break
    return m


def gen_finders():
    return {
        'H1': lambda t: re.match(r'^1\s+\S', t) is not None,
        'H2': lambda t: re.match(r'^\d\.\d\s+\S', t) is not None,
        'H3': lambda t: re.match(r'^\d+\.\d+\.\d+\s+\S', t) is not None,
        'BODY': lambda t: len(t) > 120 and not re.match(r'^\d', t) and not t.startswith('[') and not t.startswith('（'),
        'ABS': lambda t: t.startswith('摘 要：'),
        'KW': lambda t: t.startswith('关键词：'),
        'ABSEN': lambda t: t.startswith('Abstract：'),
        'KWEN': lambda t: t.startswith('Key words：'),
        'TOCT': lambda t: t.strip() == '目  录',
        'TOC1': None,
        'TOC2': None,
        'CONCT': lambda t: t.strip() == '结  论',
        'CONCB': lambda t: t.startswith('本文以'),
        'REFT': lambda t: t.strip() == '参 考 文 献',
        'REF': lambda t: t.startswith('[') and len(t) > 20,
        'ACKT': lambda t: t.strip() == '致  谢',
        'ACKB': lambda t: t.startswith('本论文是在指导老师'),
    }


def norm(s):
    """规范化便于比较：排序属性，去空白"""
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def main():
    tpl = load_tpl()
    finders = gen_finders()
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    papers = [
        ('再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx'),
    ]
    for dname, fname in papers:
        path = os.path.join(root, dname, fname)
        print('#' * 78)
        print('对比论文：', fname[:40])
        print('#' * 78)
        gd = docx.Document(path)
        for key in ['H1', 'H2', 'H3', 'BODY', 'ABS', 'KW', 'ABSEN', 'KWEN',
                    'TOCT', 'CONCT', 'CONCB', 'REFT', 'REF', 'ACKT', 'ACKB']:
            tp = tpl.get(key)
            pred = finders.get(key)
            gp = get(gd, pred) if pred else None
            if tp is None or gp is None:
                print(f'\n--- {key} ---  模板={"有" if tp else "缺"} 生成={"有" if gp else "缺"}')
                continue
            tp_ppr, tp_rpr = ppr_rpr(tp)
            gp_ppr, gp_rpr = ppr_rpr(gp)
            same_p = norm(tp_ppr) == norm(gp_ppr)
            same_r = norm(tp_rpr) == norm(gp_rpr)
            print(f'\n--- {key} ---  pPr{"一致" if same_p else "★不一致"}  rPr{"一致" if same_r else "★不一致"}')
            if not same_p:
                print('  模板 pPr:', norm(tp_ppr))
                print('  生成 pPr:', norm(gp_ppr))
            if not same_r:
                print('  模板 rPr:', norm(tp_rpr))
                print('  生成 rPr:', norm(gp_rpr))
        print()


if __name__ == '__main__':
    main()