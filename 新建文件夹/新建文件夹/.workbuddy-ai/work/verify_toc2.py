# -*- coding: utf-8 -*-
"""用 Word COM 真实分页，校验目录页码（规范化匹配，忽略点号/空格差异）"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc

W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
FILES = [
    '再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
    '装配式施工质量管理问题及优化研究——以市政项目为例',
    'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例',
]

def norm(s):
    """去掉所有空白、点号、破折号，只留汉字数字字母"""
    s = re.sub(r'[\s．.、，,]', '', s)
    return s

app = wc.gencache.EnsureDispatch('Word.Application')
app.Visible = False
app.DisplayAlerts = 0

for nm in FILES:
    path = f'{W}\\{nm}\\{nm}.docx'
    doc = app.Documents.Open(path, ReadOnly=True)
    doc.Repaginate()
    total = doc.ComputeStatistics(2)

    # 1) 建立「正文标题 -> 真实页码」
    real = {}
    for p in doc.Paragraphs:
        t = p.Range.Text.replace('\r', '').replace('\x07', '').strip()
        if not t or '\t' in t:
            continue
        pg = None
        try:
            pg = p.Range.Information(3)   # wdActiveEndPageNumber
        except Exception:
            pass
        if pg is None:
            continue
        m = re.match(r'^(\d+(?:\s*[.．]\s*\d+)*)\s*(.+)$', t)
        if m:
            key = norm(m.group(1)) + norm(m.group(2))
            real.setdefault(key, pg)
        else:
            nk = norm(t)
            if nk in ('结论', '致谢', '参考文献'):
                real.setdefault(nk, pg)

    # 2) 比对目录条目
    bad = 0
    ok = 0
    print('=' * 66)
    print(f'{nm[:26]}  总页数={total}')
    for p in doc.Paragraphs:
        raw = p.Range.Text.replace('\r', '').replace('\x07', '')
        if '\t' not in raw:
            continue
        try:
            pg_mark = p.Range.Information(3)
        except Exception:
            pg_mark = None
        if pg_mark and pg_mark > 8:   # 目录条目在目录页
            continue
        head, _, pagestr = raw.rpartition('\t')
        head = head.strip()
        m = re.match(r'^(\d+(?:\s*[.．]\s*\d+)*)\s*(.*)$', head)
        if m:
            key = norm(m.group(1)) + norm(m.group(2))
        else:
            key = norm(head)
        want = real.get(key)
        got = pagestr.strip()
        if want is None:
            print(f'  ?? {head[:26]:<28} 目录={got}  正文未找到')
            continue
        if str(want) == got:
            ok += 1
        else:
            bad += 1
            print(f'  NG {head[:26]:<28} 目录={got}  实际={want}')
    print(f'  >>> 一致 {ok} 条，不符 {bad} 条')
    doc.Close(False)

app.Quit()
