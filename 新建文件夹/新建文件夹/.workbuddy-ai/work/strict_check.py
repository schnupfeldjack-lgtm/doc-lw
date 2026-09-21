# -*- coding: utf-8 -*-
"""
按模板「明文规定」逐条实测三篇论文。
模板明文来源：炎黄职业技术学院毕业论文模板(1).docx
"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]
TW = 12700.0  # EMU per pt

def info(para):
    pf = para.paragraph_format
    runs = para.runs
    sz = fn = ea = None
    b = None
    if runs:
        r = runs[0]
        sz = r.font.size.pt if r.font.size else None
        fn = r.font.name
        b = r.font.bold
        try:
            rpr = r._element.rPr
            if rpr is not None and rpr.rFonts is not None:
                ea = rpr.rFonts.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
        except Exception:
            pass
    ls = pf.line_spacing
    if isinstance(ls, (int, float)) and ls and ls > 10000:
        ls = round(ls / TW, 1)
    li = pf.left_indent
    fi = pf.first_line_indent
    return dict(sz=sz, fn=fn, ea=ea, b=b, al=para.alignment, ls=ls,
                li=round(li/TW,1) if li else None,
                fi=round(fi/TW,1) if fi else None)

def fmt(i):
    return (f"sz={i['sz']} ea={i['ea']} b={i['b']} al={i['al']} "
            f"ls={i['ls']} li={i['li']} fi={i['fi']}")

def find(paras, pat, n=1):
    out = []
    for p in paras:
        t = p.text.strip()
        if re.match(pat, t):
            out.append(p)
            if len(out) >= n:
                break
    return out

RULES = []  # (名称, 正则, 期望描述, 校验函数)
def rule(name, pat, check):
    RULES.append((name, pat, check))

# --- 模板明文 ---
# 摘 要：（小四号宋体、加粗）
rule('摘要标题', r'^摘\s*要[:：]', lambda i: (
    i['sz'] == 12 and i['ea'] == '宋体' and i['b'] is True))
# 关键词：（小四号宋体、加粗）
rule('关键词行', r'^关键词[:：]', lambda i: (i['sz'] == 12 and i['b'] is True))
# Abstract：（小四号Times New Roman，加粗）
rule('英文摘要', r'^Abstract[:：]', lambda i: (
    str(i['fn']) == 'Times New Roman' and i['b'] is True))
rule('英文关键词', r'^Key\s*words[:：]', lambda i: (
    str(i['fn']) == 'Times New Roman' and i['b'] is True))
# 目  录 黑体小二号字加粗居中
rule('目录标题', r'^目\s*录', lambda i: (
    i['sz'] == 18 and i['ea'] == '黑体' and i['b'] is True and i['al'] == WD_ALIGN_PARAGRAPH.CENTER))
# 目录正文：宋体小四号，行距固定值28磅
rule('目录条目', r'^\d+\s*[.．]\s*\S+\t', lambda i: (
    i['sz'] == 12 and i['ea'] == '宋体' and i['ls'] == 28.0))
# 1级标题 小3号黑体加粗 段前0.5行段后0.5行
rule('一级标题', r'^\d+\s+\S', lambda i: (
    i['sz'] == 15 and i['ea'] == '黑体' and i['b'] is True))
# 2级标题 4号黑体加粗
rule('二级标题', r'^\d+[.．]\d+\s+\S', lambda i: (
    i['sz'] == 14 and i['ea'] == '黑体' and i['b'] is True))
# 3级标题 小4号黑体不加粗
rule('三级标题', r'^\d+[.．]\d+[.．]\d+\s+\S', lambda i: (
    i['sz'] == 12 and i['ea'] == '黑体' and i['b'] is not True))
# 结论 四号宋体加粗居中
rule('结论标题', r'^结\s*论$', lambda i: (
    i['sz'] == 14 and i['ea'] == '宋体' and i['b'] is True and i['al'] == WD_ALIGN_PARAGRAPH.CENTER))
# 参考文献 四号宋体加粗居中
rule('参考文献标题', r'^参\s*考\s*文\s*献$', lambda i: (
    i['sz'] == 14 and i['ea'] == '宋体' and i['b'] is True and i['al'] == WD_ALIGN_PARAGRAPH.CENTER))
# 参考文献条目 五号宋体 行距固定值22磅
rule('文献条目', r'^\[\d+\]', lambda i: (
    i['sz'] == 10.5 and i['ls'] == 22.0))
# 致谢 四号宋体加粗 左对齐
rule('致谢标题', r'^致\s*谢$', lambda i: (
    i['sz'] == 14 and i['ea'] == '宋体' and i['b'] is True and i['al'] == WD_ALIGN_PARAGRAPH.LEFT))
# 致谢正文 小四号宋体 固定值22磅
rule('致谢正文', r'^本(论文|研究|文)', lambda i: (
    i['sz'] == 12 and i['ls'] == 22.0))

for nm in FILES:
    p = f'{W}\\{nm}\\{nm}.docx'
    d = Document(p)
    paras = d.paragraphs
    print('=' * 70)
    print(nm)
    print('=' * 70)
    for name, pat, chk in RULES:
        hits = find(paras, pat, 3)
        if not hits:
            print(f'  [缺失] {name:<10} 未匹配到  pat={pat}')
            continue
        oks = [chk(info(h)) for h in hits]
        mark = 'OK ' if all(oks) else 'NG '
        i0 = info(hits[0])
        print(f'  [{mark}] {name:<10} {fmt(i0)}   << {hits[0].text[:34]}')

    # 图题/表题：5号宋体
    fig = [p2 for p2 in paras if re.match(r'^图\s*\d+[–\-—]\s*\d+', p2.text.strip())]
    tbl = [p2 for p2 in paras if re.match(r'^表\s*\d+[–\-—]\s*\d+', p2.text.strip())]
    for lab, arr in (('图题', fig), ('表题', tbl)):
        if not arr:
            print(f'  [缺失] {lab}')
            continue
        i0 = info(arr[0])
        ok = (i0['sz'] == 10.5 and i0['ea'] == '宋体')
        print(f'  [{"OK " if ok else "NG "}] {lab:<10} {fmt(i0)} 共{len(arr)}条  << {arr[0].text[:34]}')
    # 正文：小四宋体 1.5倍
    body = [p2 for p2 in paras if len(p2.text) > 120 and not p2.text.strip().startswith('[')]
    if body:
        i0 = info(body[len(body)//2])
        ok = (i0['sz'] == 12 and i0['ls'] == 1.5 and i0['fi'] == 24.0)
        print(f'  [{"OK " if ok else "NG "}] 正文段落   {fmt(i0)} 共{len(body)}段')
    # en dash 检查
    bad = [p2.text[:30] for p2 in (fig + tbl) if '–' not in p2.text[:12]]
    print(f'  [{"OK " if not bad else "NG "}] 图表编号en dash  异常{len(bad)}条 {bad[:2]}')
    print()
