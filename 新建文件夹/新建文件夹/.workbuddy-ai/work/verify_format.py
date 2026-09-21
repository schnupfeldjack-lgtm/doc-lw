# -*- coding: utf-8 -*-
"""严格校验生成论文的格式是否符合模板说明（按模板"讲到的"每一条硬编码基准）"""
import os, re
import docx
from docx.oxml.ns import qn

# ===== 模板基准（按模板说明文字硬编码）=====
MASTER = {
    'H1':   dict(size='30', ea='黑体', b=True,  jc='left',   line='360', rule='auto', before='312', after='156'),
    'H2':   dict(size='28', ea='黑体', b=True,  jc='left',   line='360', rule='auto'),
    'H3':   dict(size='24', ea='黑体', b=False, jc='left',   line='360', rule='auto'),
    'BODY': dict(size='24', ea='宋体', b=False, jc='left',   line='360', rule='auto', firstLine='570'),
    'ABS':  dict(size='24', ea='宋体', b=True,  jc='left',   line='560', rule='exact', after='72'),
    'KW':   dict(size='24', ea=None,    b=True,  jc='left',   line='360', rule='auto'),
    'ABSEN': dict(size='24', ea='Times New Roman', b=True, jc='left', line='360', rule='auto'),
    'KWEN':  dict(size='24', ea='Times New Roman', b=True, jc='left', line='360', rule='auto'),
    'TOCT': dict(size='36', ea='黑体', b=True, jc='center', line='360', rule='auto'),
    'TOC1': dict(size='24', ea='宋体', b=False, jc='left',  line='560', rule='exact'),
    'TOC2': dict(size='24', ea='宋体', b=False, jc='left',  line='560', rule='exact', firstLine='420'),
    'CONCT': dict(size='28', ea='宋体', b=True, jc='center', line='360', rule='auto'),
    'CONCB': dict(size='24', ea='宋体', b=False, jc='left',  line='440', rule='exact', firstLine='480'),
    'REFT':  dict(size='28', ea='宋体', b=True, jc='center', line='360', rule='auto'),
    'REF':   dict(size='21', ea='宋体', b=False, jc='left',  line='440', rule='exact', firstLine='573'),
    'ACKT':  dict(size='28', ea='宋体', b=True, jc='left',  line='360', rule='auto'),
    'ACKB':  dict(size='24', ea='宋体', b=False, jc='left',  line='440', rule='exact', firstLine='480'),
    'FIG':   dict(size='21', ea='宋体', b=False, jc='center'),
    'COVER': dict(size='32', ea='楷体_GB2312', b=False, jc='left'),
}


def get_para_fmt(p):
    x = p._p
    ppr = x.find(qn('w:pPr'))
    px = ppr.xml if ppr is not None else ''
    # 取第一个含文字或 tab 的 run（排除分页符、空 run）
    text_run = None
    for r in x.findall(qn('w:r')):
        if r.find(qn('w:t')) is not None or r.find(qn('w:tab')) is not None:
            text_run = r
            break
    rx = text_run.xml if text_run is not None else ''
    out = {}
    out['sz'] = (re.search(r'<w:sz w:val="(\d+)"', rx) or [None, None])[1]
    out['ea'] = (re.search(r'w:eastAsia="([^"]+)"', rx) or [None, None])[1]
    out['asc'] = (re.search(r'w:ascii="([^"]+)"', rx) or [None, None])[1]
    out['b'] = 'Y' if re.search(r'<w:b/>', rx) else '-'
    out['jc'] = (re.search(r'<w:jc w:val="([^"]+)"', px) or [None, 'left'])[1]
    out['line'] = (re.search(r'w:line="(\d+)"', px) or [None, None])[1]
    out['rule'] = (re.search(r'w:lineRule="([^"]+)"', px) or [None, None])[1]
    out['firstLine'] = (re.search(r'w:firstLine="(\d+)"', px) or [None, None])[1]
    out['before'] = (re.search(r'w:before="(\d+)"', px) or [None, None])[1]
    out['after'] = (re.search(r'w:after="(\d+)"', px) or [None, None])[1]
    return out


def match(actual, master):
    diffs = []
    for k, v in master.items():
        if v is None:
            continue
        # MASTER 用 'size'，但 get_para_fmt 输出键名是 'sz' —— 做键名映射
        if k == 'size':
            a = actual.get('sz')
        else:
            a = actual.get(k)
        if k == 'ea':
            # ea 可以为空（asc=宋体即可）；强制要求 ea=master.ea 时才比较
            if a != v:
                diffs.append(f'ea={a} (应为 {v})')
        elif k == 'b':
            expect = 'Y' if v else '-'
            if a != expect:
                diffs.append(f'b={a} (应为 {expect})')
        else:
            if a != v:
                diffs.append(f'{k}={a} (应为 {v})')
    return diffs


def short(text, n=24):
    t = text.strip().replace('\n', ' ')
    return t[:n] + ('...' if len(t) > n else '')


def check(path):
    print('=' * 78)
    print('论文：', os.path.basename(path))
    print('=' * 78)
    d = docx.Document(path)
    ps = [p for p in d.paragraphs if p.text.strip()]
    issues = []

    def find(pred):
        for p in ps:
            t = p.text.strip()
            # 排除目录条目（含 tab 的段落）和纯空白
            if '\t' in t or not t:
                continue
            if pred(t):
                return p
        return None

    checks = [
        ('H1', lambda t: re.match(r'^1\s*绪论|^1\s+\S+', t) is not None),
        ('H2', lambda t: re.match(r'^\d\.\d\s+\S', t) is not None),
        ('H3', lambda t: re.match(r'^\d+\.\d+\.\d+\s+\S', t) is not None),
        ('BODY', lambda t: len(t) > 120 and not re.match(r'^\d', t) and not t.startswith('[') and not t.startswith('（')),
        ('ABS',  lambda t: t.startswith('摘 要：')),
        ('KW',   lambda t: t.startswith('关键词：')),
        ('ABSEN', lambda t: t.startswith('Abstract：')),
        ('KWEN',  lambda t: t.startswith('Key words：')),
        ('TOCT', lambda t: t.strip() == '目  录'),
        ('TOC1', lambda t: t.startswith('1 绪论') and '\t' in t),
        ('CONCT', lambda t: t.strip() == '结  论'),
        ('CONCB', lambda t: t.startswith('本文以')),
        ('REFT',  lambda t: t.strip() == '参 考 文 献'),
        ('REF',   lambda t: t.startswith('[') and len(t) > 20),
        ('ACKT',  lambda t: t.strip() == '致  谢'),
        ('ACKB',  lambda t: t.startswith('本论文是在指导老师')),
    ]
    for name, pred in checks:
        p = find(pred)
        if p is None:
            print(f'  -- {name:6s} 未找到对应段落')
            continue
        f = get_para_fmt(p)
        m = MASTER[name]
        d_list = match(f, m)
        if d_list:
            print(f'  [BAD] {name:6s} "{short(p.text)}" -> {", ".join(d_list)}')
            issues.append(f'{name}: {", ".join(d_list)}')
        else:
            print(f'  [OK ] {name:6s} "{short(p.text)}"')
        if name == 'H1':
            # 临时调试
            print(f'    DEBUG H1: actual={ {k:v for k,v in f.items() if k in ("sz","ea","b","jc","line","rule")} }')

    # 关键词正文（第二段 run）
    pkw = find(lambda t: t.startswith('关键词：'))
    if pkw and len(pkw.runs) >= 2:
        rx = pkw.runs[1]._r.xml
        sz = (re.search(r'<w:sz w:val="(\d+)"', rx) or [None, None])[1]
        ea = (re.search(r'w:eastAsia="([^"]+)"', rx) or [None, None])[1]
        col = (re.search(r'<w:color w:val="([0-9A-Fa-f]+)"', rx) or [None, None])[1]
        ok = sz == '24' and ea == '宋体' and col == '000000'
        print(f'  [{"OK " if ok else "BAD"}] KW2    关键词正文 sz={sz} ea={ea} color={col}（模板：小四宋体）')
        if not ok:
            issues.append('关键词正文不符')

    # Key words 正文（第二段 run）
    pkwn = find(lambda t: t.startswith('Key words：'))
    if pkwn and len(pkwn.runs) >= 2:
        rx = pkwn.runs[1]._r.xml
        sz = (re.search(r'<w:sz w:val="(\d+)"', rx) or [None, None])[1]
        ea = (re.search(r'w:eastAsia="([^"]+)"', rx) or [None, None])[1]
        col = (re.search(r'<w:color w:val="([0-9A-Fa-f]+)"', rx) or [None, None])[1]
        ok = sz == '24' and ea == 'Times New Roman' and col == '000000'
        print(f'  [{"OK " if ok else "BAD"}] KWEN2  Key words正文 sz={sz} ea={ea} color={col}（模板：小四 Times New Roman）')
        if not ok:
            issues.append('Key words 正文不符')

    # 图/表题（只认居中段落）
    for p in ps:
        t = p.text.strip()
        if re.match(r'^图\d+-\d+', t) or re.match(r'^表\d+-\d+', t):
            f = get_para_fmt(p)
            if f['jc'] == 'center':
                m = MASTER['FIG']
                d_list = match(f, m)
                if d_list:
                    print(f'  [BAD] FIG    "{short(t)}" -> {", ".join(d_list)}')
                    issues.append('图/表题不符')
                else:
                    print(f'  [OK ] FIG    "{short(t)}"')

    # 全文颜色
    blue = 0
    for p in d.paragraphs:
        for r in p.runs:
            if re.search(r'<w:color w:val="(0000FF|FF0000)"', r._r.xml):
                blue += 1
    print(f'  [{"OK " if blue == 0 else "BAD"}] COLOR  非黑 run 数={blue}')
    if blue:
        issues.append(f'{blue} 个非黑 run')

    # 模板残留
    alltext = '\n'.join(p.text for p in d.paragraphs)
    residue = []
    for c in ['××××', '（可作为正文', '（本页为独立', '说明:请仔细阅读',
              '参考文献著录规则', '（同[', '（小4号宋体，1.5倍行距）',
              '毕业设计论文所列', '（五号宋体：作者', '所列出的文献',
              '题名、摘要、关键词、目录', '餐饮', '品牌', '连锁', '加强品牌']:
        if c in alltext:
            residue.append(c)
    if residue:
        print(f'  [BAD] 模板残留 {len(residue)} 项: {residue}')
        issues.append(f'模板残留: {residue}')
    else:
        print('  [OK ] 模板残留   无')

    # 封面字体（逐 cell 检查）
    cover_table = d.tables[3] if len(d.tables) >= 4 else None
    if cover_table:
        cover_cells = [
            (1, 1, '计算机系', '楷体_GB2312', '32'),
            (2, 1, '计算机应用技术', '楷体_GB2312', '32'),
        ]
        # r3c1 是题目（很长），也检查
        for ri, ci, expect_text, expect_f, expect_sz in cover_cells:
            try:
                cell = cover_table.cell(ri, ci)
                if cell.text.strip() == expect_text:
                    p = cell.paragraphs[0]
                    r = p.runs[0] if p.runs else None
                    if r:
                        rx = r._r.xml
                        ea = (re.search(r'w:eastAsia="([^"]+)"', rx) or [None, None])[1]
                        sz = (re.search(r'<w:sz w:val="(\d+)"', rx) or [None, None])[1]
                        ok = (ea == expect_f and sz == expect_sz)
                        print(f'  [{"OK " if ok else "BAD"}] COVER  r{ri}c{ci} "{expect_text}" ea={ea} sz={sz}（模板：三号楷体_GB2312）')
                        if not ok:
                            issues.append(f'封面r{ri}c{ci}字体')
            except Exception as e:
                pass
        # 题目 cell（内容很长）
        try:
            cell = cover_table.cell(3, 1)
            p = cell.paragraphs[0]
            r = p.runs[0] if p.runs else None
            if r:
                rx = r._r.xml
                ea = (re.search(r'w:eastAsia="([^"]+)"', rx) or [None, None])[1]
                sz = (re.search(r'<w:sz w:val="(\d+)"', rx) or [None, None])[1]
                ok = (ea == '楷体_GB2312' and sz == '32')
                print(f'  [{"OK " if ok else "BAD"}] COVER  r3c1 题目 ea={ea} sz={sz}（模板：三号楷体_GB2312）')
                if not ok:
                    issues.append('封面题目字体')
        except Exception:
            pass

    # 字数
    zh = sum(1 for c in '\n'.join(p.text for p in d.paragraphs) if '\u4e00' <= c <= '\u9fff')
    print(f'  [info] 中文字符总数 {zh}')

    print()
    if issues:
        print(f'!!! {os.path.basename(path)} 仍有 {len(issues)} 项问题：')
        for i in issues:
            print('   -', i)
    else:
        print(f'>>> {os.path.basename(path)} 全部格式检查通过')
    print()
    return len(issues)


if __name__ == '__main__':
    root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
    papers = [
        ('再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx'),
        ('装配式施工质量管理问题及优化研究——以市政项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例.docx'),
        ('BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx'),
    ]
    total = 0
    for d, f in papers:
        p = os.path.join(root, d, f)
        if os.path.exists(p):
            total += check(p)
        else:
            print('缺失:', p)
    print('总问题数:', total)