# -*- coding: utf-8 -*-
"""
按《炎黄职业技术学院毕业论文模板》纠正毕业论文格式（受控修改，不重建文档）。
用法: python format_thesis.py <原稿.docx> <输出.docx> <改动日志.json>
"""
import sys, re, json, copy, shutil
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.table import Table
from docx.text.paragraph import Paragraph

LOG = []

def log(zone, action, detail):
    LOG.append({'zone': zone, 'action': action, 'detail': detail})

# ---------- 基础工具 ----------
def get_or_add(el, tag):
    e = el.find(qn(tag))
    if e is None:
        e = parse_xml(f'<w:{tag.split(":")[1]} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        # keep w:pPr child order roughly: insert after first allowed position; lxml append is acceptable
        el.append(e)
    return e

def set_para_spacing(p, line=None, rule=None, before=None, after=None):
    """line/before/after 单位均为 twentieths of a point (1pt = 20)。
    1.5倍行距 => line=360, rule=auto；固定22磅 => line=440, rule=exact；段前0.5行 => before=156"""
    pPr = p._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = parse_xml('<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        pPr.insert(0, sp)
    if line is not None:
        sp.set(qn('w:line'), str(int(line)))
        sp.set(qn('w:lineRule'), 'exact' if rule == 'exact' else 'auto')
    if before is not None:
        sp.set(qn('w:before'), str(int(before)))
    if after is not None:
        sp.set(qn('w:after'), str(int(after)))

def set_para_indent(p, first=None, first_chars=None, left=None, hanging=None):
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml('<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        pPr.append(ind)
    if first is not None:
        ind.set(qn('w:firstLine'), str(int(round(first * 20))))
        if first_chars:
            ind.set(qn('w:firstLineChars'), str(first_chars))
    elif first == 0:
        ind.set(qn('w:firstLine'), '0')
    if left is not None:
        ind.set(qn('w:left'), str(int(round(left * 20))))
    if hanging is not None:
        ind.set(qn('w:hanging'), str(int(round(hanging * 20))))

def set_outline(p, level):
    pPr = p._p.get_or_add_pPr()
    ol = pPr.find(qn('w:outlineLvl'))
    if ol is None:
        ol = parse_xml('<w:outlineLvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        pPr.append(ol)
    ol.set(qn('w:val'), str(level))

def set_run(r, east=None, ascii_=None, size=None, bold=None, color=None, vert=None):
    rPr = r._r.get_or_add_rPr()
    if east or ascii_:
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            rf = parse_xml('<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.insert(0, rf)
        if east: rf.set(qn('w:eastAsia'), east)
        if ascii_:
            rf.set(qn('w:ascii'), ascii_)
            rf.set(qn('w:hAnsi'), ascii_)
    if size is not None:
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml('<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.append(sz)
        sz.set(qn('w:val'), str(int(round(size * 2))))
        szcs = rPr.find(qn('w:szCs'))
        if szcs is None:
            szcs = parse_xml('<w:szCs xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.append(szcs)
        szcs.set(qn('w:val'), str(int(round(size * 2))))
    if bold is not None:
        b = rPr.find(qn('w:b'))
        if bold and b is None:
            rPr.insert(0, parse_xml('<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        if not bold and b is not None:
            rPr.remove(b)
    if color is not None:
        c = rPr.find(qn('w:color'))
        if c is None:
            c = parse_xml('<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.append(c)
        c.set(qn('w:val'), color)

def style_runs(p, east=None, ascii_=None, size=None, bold=None, color='000000'):
    runs = p.runs
    for r in runs:
        if not r.text:
            continue  # 图片/对象 run 不改字体，避免破坏
        set_run(r, east=east, ascii_=ascii_, size=size, bold=bold, color=color)
    # 段落级 rPr（影响空段落/后续输入）
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = parse_xml('<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
        pPr.append(rPr)
    if east:
        rf = rPr.find(qn('w:rFonts'))
        if rf is None:
            rf = parse_xml('<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.insert(0, rf)
        rf.set(qn('w:eastAsia'), east)
        if ascii_:
            rf.set(qn('w:ascii'), ascii_); rf.set(qn('w:hAnsi'), ascii_)
    if size is not None:
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml('<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.append(sz)
        sz.set(qn('w:val'), str(int(round(size * 2))))
    if bold is not None:
        b = rPr.find(qn('w:b'))
        if bold and b is None:
            rPr.insert(0, parse_xml('<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        if not bold and b is not None:
            rPr.remove(b)
    if color:
        c = rPr.find(qn('w:color'))
        if c is None:
            c = parse_xml('<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
            rPr.append(c)
        c.set(qn('w:val'), color)

def para_text(p):
    return ''.join(n.text or '' for n in p._p.iter(qn('w:t')))

def has_drawing(p):
    x = p._p.xml
    return '<w:drawing>' in x or '<w:pict>' in x

def has_sectpr(p):
    pPr = p._p.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:sectPr')) is not None

# ---------- 段落分类 ----------
RE_L3 = re.compile(r'^\d+\.\d+\.\d+')
RE_L2 = re.compile(r'^\d+．\d+')
RE_L1 = re.compile(r'^\d+\s+\S')
RE_CAP = re.compile(r'^(图|表)\d+[–\-—]\d+')
RE_REF = re.compile(r'^\[\d+\]')

def classify(p):
    t = para_text(p).strip()
    if RE_L3.match(t): return 'h3'
    if RE_L2.match(t): return 'h2'
    if RE_L1.match(t) and not t.startswith('2025'): return 'h1'
    return None

# ---------- 主流程 ----------
def process(src, dst):
    shutil.copyfile(src, dst)
    doc = Document(dst)
    body = doc.element.body
    paras = [Paragraph(c, doc) for c in body.iterchildren() if c.tag == qn('w:p')]

    # --- 定位关键点 ---
    idx_abs = idx_kw = idx_absen = idx_kwen = idx_toc = None
    idx_concl = idx_ref = idx_ack = None
    first_body = None
    for i, p in enumerate(paras):
        t = para_text(p).strip()
        if t.startswith('摘 要：') and idx_abs is None: idx_abs = i
        elif t.startswith('关键词：') and idx_kw is None: idx_kw = i
        elif t.startswith('Abstract') and idx_absen is None: idx_absen = i
        elif t.startswith('Key words') and idx_kwen is None: idx_kwen = i
        elif t.replace(' ', '') == '目录' and idx_toc is None and '目' in t: idx_toc = i
        elif t.replace(' ', '') == '结论' and idx_concl is None: idx_concl = i
        elif t.replace(' ', '') == '参考文献' and idx_ref is None: idx_ref = i
        elif t.replace(' ', '') == '致谢' and idx_ack is None: idx_ack = i
        if first_body is None and idx_concl is None and classify(p) == 'h1' and i > (idx_toc or 0):
            first_body = i
    log('定位', '锚点', dict(摘要=idx_abs, 关键词=idx_kw, Abstract=idx_absen, KeyWords=idx_kwen,
                          目录=idx_toc, 正文首段=first_body, 结论=idx_concl, 参考文献=idx_ref, 致谢=idx_ack))

    # --- 1. 中文摘要 ---
    if idx_abs is not None:
        p = paras[idx_abs]
        set_para_spacing(p, line=360, rule='auto')  # 1.5 倍行距 = line 360 / auto
        set_para_indent(p, first=0)
        set_para_spacing(p, after=72)
        style_runs(p, east='宋体', size=12, bold=None)
        # 标签加粗：拆分首个 run
        runs = [r for r in p.runs if r.text]
        if runs:
            r0 = runs[0]
            txt = r0.text
            m = re.match(r'^(摘\s*要：)(.*)$', txt, re.S)
            if m:
                if not m.group(2):
                    set_run(r0, bold=True)
                else:
                    r0.text = m.group(1)
                    set_run(r0, bold=True, east='宋体', size=12)
                    newr = copy.deepcopy(r0._r)
                    # 新 run 保留原文本的剩余部分
                    for t in newr.findall(qn('w:t')):
                        t.text = m.group(2)
                    rPr = newr.find(qn('w:rPr'))
                    b = rPr.find(qn('w:b')) if rPr is not None else None
                    if b is not None: rPr.remove(b)
                    r0._r.addnext(newr)
        log('中文摘要', '行距/字体/标签加粗', '1.5倍行距(原固定28磅)、宋体小四、"摘 要："加粗、内容不加粗')

    # --- 2. 关键词 ---
    if idx_kw is not None:
        p = paras[idx_kw]
        set_para_spacing(p, line=360, rule='auto')
        set_para_indent(p, first=0)
        style_runs(p, east='宋体', size=12)
        runs = [r for r in p.runs if r.text]
        if runs:
            r0 = runs[0]
            m = re.match(r'^(关\s*键\s*词：)(.*)$', r0.text, re.S)
            if m:
                if not m.group(2):
                    set_run(r0, bold=True)
                else:
                    r0.text = m.group(1)
                    set_run(r0, bold=True, east='宋体', size=12)
                    newr = copy.deepcopy(r0._r)
                    for t in newr.findall(qn('w:t')):
                        t.text = m.group(2)
                    rPr = newr.find(qn('w:rPr'))
                    b = rPr.find(qn('w:b')) if rPr is not None else None
                    if b is not None: rPr.remove(b)
                    r0._r.addnext(newr)
        log('中文关键词', '字体/加粗', '宋体小四、1.5倍行距、"关键词："加粗、内容不加粗')

    # --- 3/4. 英文摘要与关键词 ---
    for idx, label, name in ((idx_absen, 'Abstract', '英文摘要'), (idx_kwen, 'Key words', '英文关键词')):
        if idx is None: continue
        p = paras[idx]
        set_para_spacing(p, line=360, rule='auto')
        set_para_indent(p, first=0)
        style_runs(p, east='Times New Roman', ascii_='Times New Roman', size=12)
        runs = [r for r in p.runs if r.text]
        if runs:
            r0 = runs[0]
            m = re.match(r'^(Abstract：|Key words：)(.*)$', r0.text, re.S)
            if m:
                if not m.group(2):
                    set_run(r0, bold=True)
                else:
                    r0.text = m.group(1)
                    set_run(r0, bold=True, east='Times New Roman', ascii_='Times New Roman', size=12)
                    newr = copy.deepcopy(r0._r)
                    for t in newr.findall(qn('w:t')):
                        t.text = m.group(2)
                    rPr = newr.find(qn('w:rPr'))
                    b = rPr.find(qn('w:b')) if rPr is not None else None
                    if b is not None: rPr.remove(b)
                    r0._r.addnext(newr)
        log(name, '字体/行距/加粗', 'Times New Roman 小四、1.5倍行距、标签加粗、内容不加粗')

    # --- 5. 目录：删除静态条目，插入 TOC 域 ---
    if idx_toc is not None and first_body is not None:
        p_toc = paras[idx_toc]
        p_toc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_para_spacing(p_toc, line=360, rule='auto')
        set_para_indent(p_toc, first=0)
        style_runs(p_toc, east='黑体', ascii_='黑体', size=18, bold=True)
        removed = []
        anchor = paras[first_body]._p
        for i in range(idx_toc + 1, first_body):
            p = paras[i]
            if has_sectpr(p):   # 保留分节符所在段落
                anchor = p._p
                continue
            removed.append(para_text(p)[:40])
            p._p.getparent().remove(p._p)
        # 插入 TOC 域段落（anchor 之前）
        blank = parse_xml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                          '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:p>')
        toc_p = parse_xml(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:pPr><w:spacing w:line="560" w:lineRule="exact"/></w:pPr>'
            '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr>'
            '<w:t>目录内容将在 Word 中更新生成</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            '</w:p>')
        anchor.addprevious(blank)
        anchor.addprevious(toc_p)
        log('目录', '替换为 TOC 域', f'删除静态条目 {len(removed)} 条（含手工页码）：{removed[:3]}...；'
                                     f'插入 TOC \\o "1-2" \\h \\z \\u 可更新域；标题小二黑体加粗居中')

    # --- 6. 定义 TOC 条目样式 ---
    styles_el = doc.styles.element
    for sid, name, left in (('TOC1', 'toc 1', None), ('TOC2', 'toc 2', 420), ('TOC3', 'toc 3', 840)):
        old = None
        for s in styles_el.findall(qn('w:style')):
            if s.get(qn('w:styleId')) == sid:
                old = s; break
        if old is not None:
            styles_el.remove(old)
        ind = ''
        if left:
            ind = f'<w:ind w:left="{left}" w:leftChars="{left//2 if left else 0}"/>'
        xml = (f'<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
               f'w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>'
               f'<w:basedOn w:val="Normal"/><w:qFormat/>'
               f'<w:pPr><w:spacing w:line="560" w:lineRule="exact"/>{ind}'
               f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8302"/></w:tabs></w:pPr>'
               f'<w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>'
               f'<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>')
        styles_el.append(parse_xml(xml))
    log('样式', '新增 TOC 样式', 'toc 1/2/3：宋体小四、固定28磅、右对齐制表位8302带点线、二级缩进420')

    # --- 7. 正文区 ---
    if first_body is not None:
        state = 'body'
        n_h = {'h1': 0, 'h2': 0, 'h3': 0}
        for i in range(first_body, len(paras)):
            p = paras[i]
            t = para_text(p).strip()
            # 区段切换
            if idx_concl is not None and i == idx_concl: state = 'conclusion'
            if idx_ref is not None and i == idx_ref: state = 'reference'
            if idx_ack is not None and i == idx_ack: state = 'ack'
            # 特殊标题
            if i in (idx_concl, idx_ref, idx_ack):
                name = {idx_concl: '结  论', idx_ref: '参 考 文 献', idx_ack: '致  谢'}[i]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == idx_ack else WD_ALIGN_PARAGRAPH.CENTER
                set_para_spacing(p, line=360, rule='auto')
                set_para_indent(p, first=0)
                style_runs(p, east='宋体', size=14, bold=True)
                log(name, '标题', '四号宋体加粗；' + ('致谢左对齐' if i == idx_ack else '居中'))
                continue
            kind = classify(p)
            if kind:
                n_h[kind] += 1
                lvl = {'h1': 0, 'h2': 1, 'h3': 2}[kind]
                set_outline(p, lvl)
                if kind == 'h1':
                    style_runs(p, east='黑体', ascii_='黑体', size=15, bold=True)
                    set_para_spacing(p, line=360, rule='auto', before=156, after=156)
                elif kind == 'h2':
                    style_runs(p, east='黑体', ascii_='黑体', size=14, bold=True)
                    set_para_spacing(p, line=360, rule='auto')
                    set_para_indent(p, first=0)
                else:
                    style_runs(p, east='黑体', ascii_='黑体', size=12, bold=False)
                    set_para_spacing(p, line=360, rule='auto')
                    set_para_indent(p, first=0)
                continue
            # 图表题注（居中 + 图/表编号开头）
            if RE_CAP.match(t) and p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                style_runs(p, east='宋体', size=10.5)
                set_para_spacing(p, line=440, rule='exact')
                set_para_indent(p, first=0)
                continue
            # 图片段落
            if has_drawing(p) and not t:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_para_indent(p, first=0)
                set_para_spacing(p, line=360, rule='auto')
                continue
            # 参考文献条目
            if state == 'reference' and RE_REF.match(t):
                style_runs(p, east='宋体', size=10.5)
                set_para_spacing(p, line=440, rule='exact')
                continue
            # 结论 / 致谢 内容
            if state in ('conclusion', 'ack') and t:
                style_runs(p, east='宋体', size=12)
                set_para_spacing(p, line=440, rule='exact')
                set_para_indent(p, first=24, first_chars=200)
                continue
            # 普通正文
            if t:
                style_runs(p, east='宋体', size=12)
                set_para_spacing(p, line=360, rule='auto')
                set_para_indent(p, first=24, first_chars=200)
        log('正文', '标题与正文', f'一级{n_h["h1"]}个(小三黑体加粗,段前后0.5行,outlineLvl0)、'
                                 f'二级{n_h["h2"]}个(四号黑体加粗,outlineLvl1)、三级{n_h["h3"]}个(小四黑体不加粗,outlineLvl2)；'
                                 '正文小四宋体1.5倍行距首行缩进2字符')
        log('图表', '题注', '图/表题注五号宋体、居中、固定22磅、无首行缩进；图片段落居中且去掉首行缩进')
        log('结论/参考文献/致谢', '内容', '结论与致谢小四宋体固定22磅；参考文献五号宋体固定22磅')

    # --- 8. 正文表格 ---
    tbl_count = 0
    body_started = False
    first_body_el = paras[first_body]._p if first_body is not None else None
    for child in list(body.iterchildren()):
        if child.tag == qn('w:p'):
            if first_body_el is not None and child is first_body_el:
                body_started = True
        elif child.tag == qn('w:tbl') and body_started:
            tbl = Table(child, doc)
            tbl_count += 1
            tblPr = child.find(qn('w:tblPr'))
            if tblPr is None:
                tblPr = parse_xml('<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                child.insert(0, tblPr)
            jc = tblPr.find(qn('w:jc'))
            if jc is None:
                jc = parse_xml('<w:jc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="center"/>')
                tblPr.append(jc)
            else:
                jc.set(qn('w:val'), 'center')
            for row in tbl.rows:
                for cell in row.cells:
                    for pp in cell.paragraphs:
                        for r in pp.runs:
                            if r.text:
                                set_run(r, east='宋体', size=10.5)
    if tbl_count:
        log('表格', '正文中表格', f'{tbl_count} 个：居中，单元格文字五号宋体（保留原有加粗表头）')

    doc.save(dst)
    return LOG


if __name__ == '__main__':
    src, dst, logf = sys.argv[1], sys.argv[2], sys.argv[3]
    L = process(src, dst)
    with open(logf, 'w', encoding='utf-8') as f:
        json.dump(L, f, ensure_ascii=False, indent=1)
    print('saved', dst)
    for x in L:
        print(' -', x['zone'], '|', x['action'], '|', str(x['detail'])[:100])
