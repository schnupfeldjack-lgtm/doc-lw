# -*- coding: utf-8 -*-
"""
通用论文填充脚本：
- 严格按模板母段落克隆 pPr/rPr，颜色统一为黑色（模板要求）
- 删除模板中的"说明/示例"段落，插入真实内容
- 填表、插图、生成目录（页码估算）、参考文献等
"""
import copy
import os
import re
from docx import Document
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm, Emu, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ----------------- 段落与文本工具 -----------------

def force_black(p_elem):
    """段落及所有 run 的 w:color 改为 000000"""
    for col in p_elem.iter(qn('w:color')):
        col.set(qn('w:val'), '000000')


def blacken_rpr(rpr):
    """对单个 rPr 元素（深拷贝过来的）改色"""
    if rpr is None:
        return
    for col in rpr.findall(qn('w:color')):
        col.set(qn('w:val'), '000000')
    # 若 rPr 没有 color 元素，添加一个确保黑色
    if rpr.find(qn('w:color')) is None:
        col = OxmlElement('w:color')
        col.set(qn('w:val'), '000000')
        rpr.append(col)


def force_run_format(para, font, size):
    """强制把段落内所有 run 的字体和字号设为指定值（覆盖继承）。

    说明：模板母段落的 XML 与其自身写明的规范并不完全一致（例如模板正文母段落
    没有写 eastAsia 字体、英文摘要母段落写成了三号黑体、结论/参考文献标题母段落
    没有加粗）。这里以"模板写明的规范"为准做强制统一，保证成品符合模板规范。"""
    for r in para.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            r.insert(0, rpr)
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.insert(0, rfonts)
        rfonts.set(qn('w:eastAsia'), font)
        rfonts.set(qn('w:ascii'), font)
        rfonts.set(qn('w:hAnsi'), font)
        sz = rpr.find(qn('w:sz'))
        if sz is None:
            sz = OxmlElement('w:sz')
            rpr.append(sz)
        sz.set(qn('w:val'), size)
        szcs = rpr.find(qn('w:szCs'))
        if szcs is None:
            szcs = OxmlElement('w:szCs')
            rpr.append(szcs)
        szcs.set(qn('w:val'), size)
        col = rpr.find(qn('w:color'))
        if col is None:
            col = OxmlElement('w:color')
            rpr.append(col)
        col.set(qn('w:val'), '000000')


def force_para_spacing(para, line, rule):
    """强制段落行距（用于结论/致谢正文等行距为固定值22磅的位置）"""
    ppr = para.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        para.insert(0, ppr)
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        ppr.append(sp)
    sp.set(qn('w:line'), str(line))
    sp.set(qn('w:lineRule'), rule)


def ensure_h1_format(para):
    """章标题：小3号(30) 黑体加粗，段前0.5行段后0.5行"""
    force_run_format(para, '黑体', '30')
    ppr = para.find(qn('w:pPr'))
    if ppr is not None:
        sp = ppr.find(qn('w:spacing'))
        if sp is not None:
            sp.set(qn('w:before'), '312')
            sp.set(qn('w:after'), '156')
        for r in para.findall(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is not None and rpr.find(qn('w:b')) is None:
                rpr.append(OxmlElement('w:b'))


def ensure_h2_format(para):
    """二级标题：完全沿用模板母段落 rPr（4号黑体加粗），不做覆盖"""
    force_run_format(para, '黑体', '28')
    for r in para.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is not None and rpr.find(qn('w:b')) is None:
            rpr.append(OxmlElement('w:b'))


def ensure_h3_format(para):
    """三级标题：完全沿用模板母段落 rPr（小4号黑体，不加粗），不做覆盖"""
    force_run_format(para, '黑体', '24')


def ensure_title_format(para, font='黑体', size='30', bold=True):
    """结论/参考文献/致谢标题：完全沿用模板母段落 rPr，不做覆盖
       （模板母段落 run 的 rPr 就是最终标准，强行加粗会与模板不一致）"""
    force_run_format(para, font, size)
    for r in para.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            r.insert(0, rpr)
        if bold and rpr.find(qn('w:b')) is None:
            rpr.append(OxmlElement('w:b'))


def make_kw_para(src_p, label, content, content_font='宋体'):
    """关键词段落：标签按母段落样式（P40 黑体加粗小四 / P46 加粗小四），
       内容用 content_font 小四不加粗（中文=宋体，英文=Times New Roman）。
       符合模板：「关键词：（小四号宋体、加粗）  ×××（小四号宋体）」
                 「Key words：（小四号Times New Roman、加粗）×××（小四号Times New Roman）」"""
    new_p = copy.deepcopy(src_p)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in new_p.findall(qn(tag)):
            new_p.remove(e)
    force_black(new_p)
    rpr_label = get_rpr(src_p)
    blacken_rpr(rpr_label)
    r1 = OxmlElement('w:r')
    if rpr_label is not None:
        r1.append(copy.deepcopy(rpr_label))
    t1 = OxmlElement('w:t')
    t1.text = label
    t1.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r1.append(t1)
    new_p.append(r1)
    # 内容 run：content_font 小四不加粗
    rpr_text = OxmlElement('w:rPr')
    rfonts = OxmlElement('w:rFonts')
    rfonts.set(qn('w:eastAsia'), content_font)
    if content_font == '宋体':
        rfonts.set(qn('w:ascii'), 'Times New Roman')
        rfonts.set(qn('w:hAnsi'), 'Times New Roman')
    else:
        rfonts.set(qn('w:ascii'), content_font)
        rfonts.set(qn('w:hAnsi'), content_font)
    rpr_text.append(rfonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '24')
    rpr_text.append(sz)
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), '24')
    rpr_text.append(szcs)
    col = OxmlElement('w:color')
    col.set(qn('w:val'), '000000')
    rpr_text.append(col)
    r2 = OxmlElement('w:r')
    r2.append(rpr_text)
    t2 = OxmlElement('w:t')
    t2.text = '  ' + content
    t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r2.append(t2)
    new_p.append(r2)
    force_black(new_p)
    return new_p


def get_rpr(src_p):
    """取母段落中"带文字(w:t)"的 run 的 rPr（深拷贝）——这才是正文文字真正使用的格式。
       注意：模板有些段落第一个 run 没有 w:t（如分页符/书签 run），
       直接取"第一个 run"会取到错误的格式（例如章标题取到 sz=20 而非 30）。"""
    # 1) 优先取带 w:t 的 run
    for r in src_p.findall(qn('w:r')):
        if r.find(qn('w:t')) is not None:
            rpr = r.find(qn('w:rPr'))
            if rpr is not None:
                return copy.deepcopy(rpr)
    # 2) 退回第一个有 rPr 的 run
    for r in src_p.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is not None:
            return copy.deepcopy(rpr)
    # 3) 用段落级 rPr
    ppr = src_p.find(qn('w:pPr'))
    if ppr is not None:
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None:
            return copy.deepcopy(rpr)
    return None


def make_p_with_text(src_p, text):
    """克隆 src_p 的 pPr，清空所有 run/hyperlink，加一个新 run 装 text；颜色改黑。"""
    new_p = copy.deepcopy(src_p)
    # 移除 run 与 hyperlink
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in new_p.findall(qn(tag)):
            new_p.remove(e)
    force_black(new_p)
    if text:
        rpr = get_rpr(src_p)
        # 强制把深拷贝来的 rpr 也改黑（模板示例段落的 rpr 多为蓝色 0000FF）
        blacken_rpr(rpr)
        r = OxmlElement('w:r')
        if rpr is not None:
            r.append(rpr)
        t = OxmlElement('w:t')
        t.text = text
        if text.startswith(' ') or text.endswith(' ') or '  ' in text or '\t' in text:
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        new_p.append(r)
        # 再次保险：append 后把段落整体再过一遍改色
        force_black(new_p)
    return new_p


def make_p_picture(src_p, img_path, caption_text, width_cm=13.5):
    """插入图片段落 + 图题段落（保留模板母段落格式）"""
    p_img = copy.deepcopy(src_p)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in p_img.findall(qn(tag)):
            p_img.remove(e)
    force_black(p_img)
    # 居中
    ppr = p_img.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        p_img.insert(0, ppr)
    jc = ppr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        ppr.append(jc)
    jc.set(qn('w:val'), 'center')

    rpr_pic = OxmlElement('w:rPr')
    r = OxmlElement('w:r')
    r.append(rpr_pic)
    drawing = OxmlElement('w:drawing')
    inline = etree.SubElement(drawing, qn('wp:inline'),
                              nsmap={'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'})
    etree.SubElement(inline, qn('wp:extent')).set('cx', str(int(width_cm * 360000)))
    etree.SubElement(inline, qn('wp:extent')).set('cy', '0')  # 比例由 pic 自动
    # 简化的 picture xml，依赖 python-docx 后处理
    r.append(drawing)
    p_img.append(r)
    return p_img


def make_toc_para(src_p, title, page):
    """克隆目录母段落，构造"标题\\t 页码"形式，保留前导点 tab。"""
    new_p = copy.deepcopy(src_p)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in new_p.findall(qn(tag)):
            new_p.remove(e)
    force_black(new_p)
    rpr = get_rpr(src_p)
    blacken_rpr(rpr)
    # 标题 run
    r1 = OxmlElement('w:r')
    if rpr is not None:
        r1.append(copy.deepcopy(rpr))
    t1 = OxmlElement('w:t')
    t1.text = title
    t1.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r1.append(t1)
    new_p.append(r1)
    # tab run
    r2 = OxmlElement('w:r')
    if rpr is not None:
        r2.append(copy.deepcopy(rpr))
    tab = OxmlElement('w:tab')
    r2.append(tab)
    new_p.append(r2)
    # 页码 run
    r3 = OxmlElement('w:r')
    if rpr is not None:
        r3.append(copy.deepcopy(rpr))
    t3 = OxmlElement('w:t')
    t3.text = str(page)
    r3.append(t3)
    new_p.append(r3)
    force_black(new_p)
    return new_p


def insert_after(anchor_elem, new_elem):
    anchor_elem.addnext(new_elem)
    return new_elem


def insert_batch(anchor_elem, new_elems):
    """顺序插入到 anchor 之后，返回最后一个新元素"""
    last = anchor_elem
    for e in new_elems:
        last.addnext(e)
        last = e
    return last


def remove_p(p):
    pe = p._p if hasattr(p, '_p') else p
    parent = pe.getparent()
    parent.remove(pe)


# ----------------- 解析 markdown -----------------

def parse_md(path):
    """解析正文 md，返回段落列表，每项 dict: {type, text, level, caption_key}"""
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    # 去除摘要/关键词块（已在 meta 中）
    text = re.sub(r'@@ABSTRACT_CN@@.*?(?=@@BODY@@)', '', text, flags=re.S)
    text = text.replace('@@BODY@@', '')
    paragraphs = []
    mode = None  # None | 'conclusion' | 'ack'
    for line in text.splitlines():
        s = line.rstrip()
        if not s:
            continue
        # 结论 / 致谢 区块：标记之后的所有行归入对应类型
        if s.startswith('@@CONCLUSION@@'):
            mode = 'conclusion'
            continue
        if s.startswith('@@ACK@@'):
            mode = 'ack'
            continue
        if s.startswith('## '):
            paragraphs.append({'type': 'h1', 'text': s[3:].strip()})
        elif s.startswith('### '):
            paragraphs.append({'type': 'h2', 'text': s[4:].strip()})
        elif s.startswith('#### '):
            paragraphs.append({'type': 'h3', 'text': s[5:].strip()})
        elif s.startswith('[FIG:'):
            m = re.match(r'\[FIG:([^:]+):(.+)\]', s)
            if m:
                paragraphs.append({'type': 'fig', 'key': m.group(1), 'caption': m.group(2)})
        elif s.startswith('[TBL:'):
            m = re.match(r'\[TBL:([^:]+):(.+)\]', s)
            if m:
                paragraphs.append({'type': 'tbl', 'key': m.group(1), 'caption': m.group(2)})
        else:
            if mode == 'conclusion':
                paragraphs.append({'type': 'conclusion', 'text': s})
            elif mode == 'ack':
                paragraphs.append({'type': 'ack', 'text': s})
            else:
                paragraphs.append({'type': 'p', 'text': s})
    return paragraphs


def parse_meta(path):
    """解析 meta md，提取各标记块"""
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    out = {}
    pattern = re.compile(r'@@(\w+)@@\n(.*?)(?=@@\w+@@|\Z)', re.S)
    for m in pattern.finditer(text):
        key = m.group(1)
        val = m.group(2).strip()
        out[key] = val
    return out


# ----------------- 填充 docx -----------------

def fill(paper_dir, docx_name, title,
         md_body_files, md_meta_file,
         figures,  # dict: {fig_key: image_path}
         table_html=None,  # 预留：自定义表格（Word 原生表）
         target_chars_per_page=1250):
    src_path = os.path.join(paper_dir, docx_name)
    doc = Document(src_path)
    body = doc.element.body

    # 收集 blocks（段落+表格）
    blocks = []
    for child in list(body.iterchildren()):
        if child.tag == qn('w:p'):
            blocks.append(('p', child))
        elif child.tag == qn('w:tbl'):
            blocks.append(('t', child))

    # 母段落（在原始模板中按索引找；此处 src 是已复制但内容未改的模板）
    def find_p_by_text(start_text):
        for kind, e in blocks:
            if kind != 'p':
                continue
            text = ''.join(e.itertext()).strip()
            if text.startswith(start_text):
                return e
        return None

    M_H1 = find_p_by_text('1  引言')           # 章标题
    M_BODY = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '小4号宋体，1.5倍行距' in text:
            M_BODY = e
            break
    M_H2 = find_p_by_text('1．1')
    if M_H2 is None:
        for kind, e in blocks:
            if kind != 'p':
                continue
            text = ''.join(e.itertext())
            if '正文2级标题' in text:
                M_H2 = e
                break
    M_H3 = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '正文3级标题' in text:
            M_H3 = e
            break
    M_ABS = find_p_by_text('摘 要：')
    M_KW = find_p_by_text('关键词：')
    M_ABSTRACT_EN = find_p_by_text('Abstract：')
    M_KEYWORDS_EN = find_p_by_text('Key words：')
    M_TOC_TITLE = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '目  录' in text and '黑体小二号' in text:
            M_TOC_TITLE = e
            break
    M_TOC1 = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if text.startswith('1.引言') and '\t' in text:
            M_TOC1 = e
            break
    M_TOC2 = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if text.startswith('4.1') and '\t' in text:
            M_TOC2 = e
            break
    M_CONC_TITLE = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if text.strip().startswith('结  论') and '四号宋体' in text:
            M_CONC_TITLE = e
            break
    M_CONC_BODY = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '小四号宋体，行距为固定值22磅' in text:
            M_CONC_BODY = e
            break
    M_REF_TITLE = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '参 考 文 献' in text and '四号宋体' in text:
            M_REF_TITLE = e
            break
    M_REF = find_p_by_text('[1]')
    M_ACK_TITLE = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if text.strip().startswith('致  谢'):
            M_ACK_TITLE = e
            break
    M_ACK_BODY = None
    for kind, e in blocks:
        if kind != 'p':
            continue
        text = ''.join(e.itertext())
        if '小四号宋体，行距为固定值22磅' in text and e is not M_CONC_BODY:
            M_ACK_BODY = e
            break

    # ============ 解析正文 & meta ============
    body_paras = []
    for f in md_body_files:
        body_paras.extend(parse_md(f))
    meta = parse_meta(md_meta_file)

    # 关键修复：摘要 / 关键词 / Abstract / Key words / 结论 / 致谢 写在正文文件(part1~part3)
    # 的 @@XXX@@ 块中，而 parse_md() 会把这些块整段删掉（只保留正文），
    # 若此处只从 meta 文件读取，摘要就会是空的（只剩"摘 要："标签）。
    # 因此把正文文件里解析到的 @@ 块合并进 meta（meta 中已有的不覆盖）。
    for _f in md_body_files:
        try:
            _m = parse_meta(_f)
        except Exception:
            continue
        for _k in ('ABSTRACT_CN', 'KEYWORDS_CN', 'ABSTRACT_EN', 'KEYWORDS_EN',
                   'CONCLUSION', 'ACK'):
            if _k in _m and not meta.get(_k):
                meta[_k] = _m[_k]

    # 结论 / 致谢正文：优先取 meta，否则从正文解析结果里提取
    # （@@CONCLUSION@@ / @@ACK@@ 标记后的行会被 parse_md 标记为对应类型）
    conc_text = meta.get('CONCLUSION', '')
    if not conc_text:
        conc_text = '\n'.join(item.get('text', '')
                              for item in body_paras if item['type'] == 'conclusion')
    ack_text = meta.get('ACK', '')
    if not ack_text:
        ack_text = '\n'.join(item.get('text', '')
                             for item in body_paras if item['type'] == 'ack')

    # ============ 处理开题报告表（TABLE 0） ============
    tbls = [e for kind, e in blocks if kind == 't']
    T0 = tbls[0]
    for ri, row in enumerate(T0.findall(qn('w:tr'))):
        cells = row.findall(qn('w:tc'))
        seen = set()
        for ci, cell in enumerate(cells):
            cell_id = id(cell)
            if cell_id in seen:
                continue
            seen.add(cell_id)
            ps_in = cell.findall(qn('w:p'))
            texts = [''.join(p.itertext()).strip() for p in ps_in]
            # 根据行处理
            if ri == 1 and ci == 0:
                # 课题名称
                if ps_in:
                    for p in ps_in[1:]:
                        remove_p(p)
                    set_p_runs_text(ps_in[0], meta.get('KT_TITLE', title))
            elif ri == 2 and ci == 0:
                # 课题概况
                clear_cell(cell)
                lines = [
                    ('课题概况（主旨，思路，方法）', True),
                    ('　　' + meta.get('KT_ZHUIZHI', ''), False),
                    ('　　' + meta.get('KT_SILU', ''), False),
                    ('　　' + meta.get('KT_FANGFA', ''), False),
                ]
                for text, is_label in lines:
                    add_cell_para(cell, text, is_label)
            elif ri == 3 and ci == 0:
                # 导师意见
                clear_cell(cell)
                add_cell_para(cell, '导师意见：', True)
                add_cell_para(cell, '　　' + meta.get('KT_DAOSHI', ''))
                add_cell_para(cell, '')
                add_cell_para(cell, '')
                add_cell_para(cell, '签名：____________　时间：2025.9.10')
            elif ri == 4 and ci == 0:
                # 评审组长意见
                clear_cell(cell)
                add_cell_para(cell, '论文评审组长意见：', True)
                add_cell_para(cell, '　　' + meta.get('KT_ZUZHANG', ''))
                add_cell_para(cell, '')
                add_cell_para(cell, '签名：____________　时间：2025.9.12')

    # ============ 处理任务书表（TABLE 1） ============
    T1 = tbls[1]
    for ri, row in enumerate(T1.findall(qn('w:tr'))):
        cells = row.findall(qn('w:tc'))
        for ci, cell in enumerate(cells):
            if ci > 0:
                continue
            if ri == 0:
                clear_cell(cell)
                add_cell_para(cell, '课题名称：' + meta.get('KT_TITLE', title))
            elif ri == 1:
                clear_cell(cell)
                add_cell_para(cell, '本课题主要内容：', True)
                for line in meta.get('RWS_CONTENT', '').split('\n'):
                    add_cell_para(cell, line)
            elif ri == 2:
                clear_cell(cell)
                add_cell_para(cell, '本课题研究方法与要求：', True)
                add_cell_para(cell, '')
                add_cell_para(cell, '方法：' + meta.get('RWS_FANGFA', ''))
                add_cell_para(cell, '')
                add_cell_para(cell, '要求：' + meta.get('RWS_YAOQIU', ''))
            elif ri == 3:
                clear_cell(cell)
                add_cell_para(cell, '本课题研究进度安排：', True)
                for line in meta.get('RWS_JINDU', '').split('\n'):
                    add_cell_para(cell, line)
            elif ri == 4:
                clear_cell(cell)
                add_cell_para(cell, '主要参考文献：', True)
                for line in meta.get('RWS_CANKAO', '').split('\n'):
                    if line.strip():
                        add_cell_para(cell, line)

    # ============ 处理中期检查表（TABLE 2） ============
    T2 = tbls[2]
    stage_keys = ['ZQ_S1', 'ZQ_S2', 'ZQ_S3', 'ZQ_S4']
    rows2 = T2.findall(qn('w:tr'))
    for ri in (2, 3, 4, 5):
        if ri >= len(rows2):
            continue
        cells = rows2[ri].findall(qn('w:tc'))
        seen = set()
        for ci, cell in enumerate(cells):
            cid = id(cell)
            if cid in seen:
                continue
            seen.add(cid)
            if ci != 1:
                continue
            stage_text = meta.get(stage_keys[ri - 2], '')
            lines = stage_text.split('\n')
            clear_cell(cell)
            for j, ln in enumerate(lines):
                add_cell_para(cell, ln, bold=(j == 0))

    # ============ 处理封面表（TABLE 3） ============
    T3 = tbls[3]
    rows3 = T3.findall(qn('w:tr'))
    for ri, row in enumerate(rows3):
        cells = row.findall(qn('w:tc'))
        seen = set()
        for ci, cell in enumerate(cells):
            cid = id(cell)
            if cid in seen:
                continue
            seen.add(cid)
            if ri == 1 and ci == 1:
                clear_cell(cell)
                add_cell_para(cell, '计算机系', font='楷体_GB2312', size='32')
            elif ri == 2 and ci == 1:
                clear_cell(cell)
                add_cell_para(cell, '计算机应用技术', font='楷体_GB2312', size='32')
            elif ri == 3 and ci == 1:
                clear_cell(cell)
                add_cell_para(cell, title, font='楷体_GB2312', size='32')

    # ============ 处理年月表（TABLE 4） ============
    T4 = tbls[4]
    rows4 = T4.findall(qn('w:tr'))
    for row in rows4:
        cells = row.findall(qn('w:tc'))
        seen = set()
        for ci, cell in enumerate(cells):
            cid = id(cell)
            if cid in seen:
                continue
            seen.add(cid)
            if ci == 0:
                clear_cell(cell)
                add_cell_para(cell, '2025')
            elif ci == 2:
                clear_cell(cell)
                add_cell_para(cell, '9')

    # ============ 处理摘要 / 关键词 ============
    abs_cn = meta.get('ABSTRACT_CN', '')
    kw_cn = meta.get('KEYWORDS_CN', '')
    abs_en = meta.get('ABSTRACT_EN', '')
    kw_en = meta.get('KEYWORDS_EN', '')

    # 中文摘要（模板：小四号宋体、加粗）
    abs_para = make_p_with_text(M_ABS, '摘 要：' + abs_cn)
    force_run_format(abs_para, '宋体', '24')
    M_ABS.addprevious(abs_para)
    remove_p(M_ABS)

    # 中文关键词
    kw_para = make_kw_para(M_KW, '关键词：', kw_cn, '宋体')
    force_run_format(kw_para, '宋体', '24')  # 模板明文：关键词标签与内容均为小四号宋体
    M_KW.addprevious(kw_para)
    remove_p(M_KW)

    # Abstract（模板：小四号 Times New Roman、加粗）
    ab_para = make_p_with_text(M_ABSTRACT_EN, 'Abstract：' + abs_en)
    force_run_format(ab_para, 'Times New Roman', '24')
    M_ABSTRACT_EN.addprevious(ab_para)
    remove_p(M_ABSTRACT_EN)

    # Key words（模板：小四号 Times New Roman，加粗）
    kw2_para = make_kw_para(M_KEYWORDS_EN, 'Key words：', kw_en, 'Times New Roman')
    force_run_format(kw2_para, 'Times New Roman', '24')
    M_KEYWORDS_EN.addprevious(kw2_para)
    remove_p(M_KEYWORDS_EN)

    # ============ 目录 ============
    # 先估算每章起始页码
    page = 1
    toc_entries = []
    for item in body_paras:
        if item['type'] == 'h1':
            toc_entries.append({'level': 1, 'title': item['text'], 'page': page})
        elif item['type'] == 'h2':
            toc_entries.append({'level': 2, 'title': item['text'], 'page': page})
        # 占位估算
        if item['type'] == 'h1':
            page += 3
        elif item['type'] == 'h2':
            page += 1
        elif item['type'] == 'p':
            page += max(1, len(item['text']) // target_chars_per_page + 1)
        elif item['type'] == 'fig':
            page += 1
        elif item['type'] == 'tbl':
            page += 1
    # 结论 / 致谢 / 参考文献
    conclusion_text = meta.get('CONCLUSION', '')
    toc_entries.append({'level': 1, 'title': '结论', 'page': page})
    page += max(1, len(conclusion_text) // target_chars_per_page + 1)
    toc_entries.append({'level': 1, 'title': '致谢', 'page': page})
    # 注意：此处不要重新赋值 ack_text（会覆盖前面从正文提取到的致谢文本，导致致谢正文为空）
    page += max(1, len(ack_text) // target_chars_per_page + 1)
    toc_entries.append({'level': 1, 'title': '参考文献', 'page': page})

    # 构造新目录：在原 M_TOC1（模板示例条目）位置之前插入
    toc_title_para = make_p_with_text(M_TOC_TITLE, '目  录')
    M_TOC_TITLE.addprevious(toc_title_para)
    remove_p(M_TOC_TITLE)

    # 找到第一个目录示例段落作为锚点
    toc_anchor = M_TOC1
    new_toc_paras = []
    for e in toc_entries:
        if e['level'] == 1:
            p = make_toc_para(M_TOC1, e['title'], e['page'])
        else:
            p = make_toc_para(M_TOC2, e['title'], e['page'])
        force_run_format(p, '宋体', '24')
        new_toc_paras.append(p)
    insert_batch(toc_anchor, new_toc_paras)
    # 删除原模板所有目录示例（保留新插入的目录条目，跳过空段以便 M_H1 前留空行）
    toc_keep = set(id(e) for e in new_toc_paras)
    cur = toc_anchor
    while cur is not None and cur is not M_H1:
        nxt = cur.getnext()
        if cur.tag == qn('w:p') and id(cur) not in toc_keep:
            text = ''.join(cur.itertext()).strip()
            if text:
                remove_p(cur)
        cur = nxt
    # 清理目录标题与目录条目之间的模板说明段落
    cur = toc_title_para.getnext()
    while cur is not None and cur is not toc_anchor:
        nxt = cur.getnext()
        if cur.tag == qn('w:p'):
            text = ''.join(cur.itertext()).strip()
            if text and ('（空' in text or '目录正文部分' in text or '首行前空' in text):
                remove_p(cur)
        cur = nxt
    # 模板要求：目录"首行前空1行"—— 插入一个空段
    blank_para = copy.deepcopy(M_TOC1)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in blank_para.findall(qn(tag)):
            blank_para.remove(e)
    force_black(blank_para)
    toc_title_para.addnext(blank_para)

    # ============ 正文 ============
    body_anchor = M_H1  # 章标题母段落位置作为锚点
    new_body = []
    # 章节计数（图编号）
    chapter = 0
    figure_seq = 0
    for item in body_paras:
        t = item['type']
        if t == 'h1':
            chapter += 1
            figure_seq = 0
            txt = item['text']
            new_body.append(make_p_with_text(M_H1, txt))
            # 模板要求：章标题小3号黑体加粗，段前0.5行段后0.5行
            ensure_h1_format(new_body[-1])
        elif t == 'h2':
            new_body.append(make_p_with_text(M_H2, item['text']))
            ensure_h2_format(new_body[-1])
        elif t == 'h3':
            new_body.append(make_p_with_text(M_H3, item['text']))
            ensure_h3_format(new_body[-1])
        elif t == 'p':
            # 按段落处理：插入段落（模板：小4号宋体，1.5倍行距）
            for line in split_long_para(item['text']):
                np = make_p_with_text(M_BODY, line)
                force_run_format(np, '宋体', '24')
                new_body.append(np)
        elif t == 'fig':
            figure_seq += 1
            cap = item['caption']
            # 调整 caption 中的"图X-Y"
            cap = re.sub(r'图\s*\d+-\d+', f'图{chapter}–{figure_seq}', cap, count=1)
            # 找对应图片
            img_path = figures.get(item['key'])
            if img_path and os.path.exists(img_path):
                # 图片段落
                p_pic = make_p_picture_paragraph(M_BODY, img_path)
                new_body.append(p_pic)
            # 图题段落（模板：图与表字体一律5号宋体）
            cap_para = make_caption_paragraph(M_REF, cap)
            force_run_format(cap_para, '宋体', '21')
            new_body.append(cap_para)
        elif t == 'tbl':
            cap = item['caption']
            cap = re.sub(r'表\s*\d+-\d+', f'表{chapter}–1', cap, count=1)
            cap_para = make_caption_paragraph(M_REF, cap)
            force_run_format(cap_para, '宋体', '21')
            new_body.append(cap_para)
            # 插入内置 Word 表格
            tbl_elem = build_eval_table_docx(item.get('key'), title)
            new_body.append(tbl_elem)
        elif t == 'conclusion':
            pass  # 后面单独处理
        elif t == 'ack':
            pass

    insert_batch(body_anchor, new_body)
    # 删除原模板所有正文示例段落与注释段落（用模板独有标记识别）
    TEMPLATE_BODY_MARKERS = ['××××', '（可作为正文', '（小4号', '（小四号宋体，行距为固定值22磅',
                             '（本页为独立', '（每个项目', '注：1．正文中',
                             '正文各页的格式', '为保证打印', '（作为正文']
    stop = M_CONC_TITLE
    cur = body_anchor
    while cur is not None and cur is not stop:
        nxt = cur.getnext()
        if cur.tag == qn('w:p'):
            text = ''.join(cur.itertext())
            if any(m in text for m in TEMPLATE_BODY_MARKERS):
                remove_p(cur)
        cur = nxt
        if cur is None:
            break

    # ============ 结论 ============
    conc_paras = split_long_para(conc_text)
    conc_title_para = make_p_with_text(M_CONC_TITLE, '结  论')
    # 模板要求：结论"本页为独立页"—— 在标题文字后追加分页符（不影响文字 run 的字体格式）
    br_run = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    br_run.append(br)
    # 找到标题文字的 run，附加在它之后（不是 pPr 之后，避免跑到文字 run 前面把标题变成第二个 run）
    text_run = None
    for _r in conc_title_para.findall(qn('w:r')):
        if _r.find(qn('w:t')) is not None:
            text_run = _r
            break
    if text_run is not None:
        text_run.addnext(br_run)
    else:
        conc_title_para.append(br_run)
    insert_after(M_CONC_TITLE, conc_title_para)
    ensure_title_format(conc_title_para, '宋体', '28')
    remove_p(M_CONC_TITLE)
    conc_body_paras = [make_p_with_text(M_CONC_BODY, t) for t in conc_paras]
    # 模板：结论正文小四号宋体，行距为固定值22磅
    for _p in conc_body_paras:
        force_run_format(_p, '宋体', '24')
        force_para_spacing(_p, 440, 'exact')
        # 首行缩进 2 字符
        ppr = _p.find(qn('w:pPr'))
        if ppr is not None:
            ind = ppr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                ppr.append(ind)
            ind.set(qn('w:firstLine'), '480')
            ind.set(qn('w:firstLineChars'), '200')
    insert_batch(M_CONC_BODY, conc_body_paras)
    remove_p(M_CONC_BODY)
    # 清理结论与参考文献标题之间的模板说明段落（保留新结论标题与正文）
    conc_keep = set(id(e) for e in conc_body_paras)
    conc_keep.add(id(conc_title_para))
    cur = M_REF_TITLE.getprevious()
    while cur is not None and cur is not conc_title_para:
        prev = cur.getprevious()
        if cur.tag == qn('w:p') and id(cur) not in conc_keep:
            text = ''.join(cur.itertext())
            if '××××' in text or '说明：结论' in text or '说明:结论' in text \
                    or text.strip().startswith('（空') or '可以在结论或讨论' in text:
                remove_p(cur)
        cur = prev

    # ============ 参考文献 ============
    refs_text = meta.get('REFS', '')
    refs_list = [l for l in refs_text.split('\n') if l.strip()]
    ref_title_para = make_p_with_text(M_REF_TITLE, '参 考 文 献')
    insert_after(M_REF_TITLE, ref_title_para)
    ensure_title_format(ref_title_para, '宋体', '28')
    remove_p(M_REF_TITLE)
    ref_paras = [make_p_with_text(M_REF, l) for l in refs_list]
    # 模板：参考文献条目五号宋体，行距为固定值22磅
    for _p in ref_paras:
        force_run_format(_p, '宋体', '21')
    insert_batch(M_REF, ref_paras)
    # 删除原模板参考文献示例与说明（用模板独有标记识别，跳过新插入的条目）
    TEMPLATE_REF_MARKERS = ['××××', '（五号宋体：作者', '（同[', '说明:请仔细阅读',
                             '参考文献著录规则', '所列出的文献', '参考文献的角注',
                             '角注的数字序号', '题名、摘要', '2.毕业设计论文所列',
                             '3.参考文献的角注', '在正文中，应用文献']
    ref_keep = set(id(e) for e in ref_paras)
    cur = M_REF
    while cur is not None and cur is not M_ACK_TITLE:
        nxt = cur.getnext()
        if cur.tag == qn('w:p') and id(cur) not in ref_keep:
            text = ''.join(cur.itertext()).strip()
            if any(m in text for m in TEMPLATE_REF_MARKERS) or text.startswith('[1]') \
                    or text.startswith('例如') or text in ('…………', '（空2行）'):
                remove_p(cur)
        cur = nxt

    # ============ 致谢 ============
    ack_paras = split_long_para(ack_text)
    ack_title_para = make_p_with_text(M_ACK_TITLE, '致  谢')
    insert_after(M_ACK_TITLE, ack_title_para)
    ensure_title_format(ack_title_para, '宋体', '28')
    remove_p(M_ACK_TITLE)
    ack_new = [make_p_with_text(M_ACK_BODY, t) for t in ack_paras]
    # 模板：致谢正文小四号宋体，行距为固定值22磅
    for _p in ack_new:
        force_run_format(_p, '宋体', '24')
        force_para_spacing(_p, 440, 'exact')
        ppr = _p.find(qn('w:pPr'))
        if ppr is not None:
            ind = ppr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                ppr.append(ind)
            ind.set(qn('w:firstLine'), '480')
            ind.set(qn('w:firstLineChars'), '200')
    insert_batch(M_ACK_BODY, ack_new)
    remove_p(M_ACK_BODY)

    # ============ 全局兜底：清除模板示例/说明残留 ============
    # 模板要求：各类说明（蓝色/红色字体表示）在参阅后自行删除。
    # 这里对整个文档再扫一遍，凡命中模板占位/说明特征的段落一律删除，避免漏网。
    RESIDUE_MARKERS = [
        '××××', '×××', 'Maham', 'Curzon', '…………',
        '3.参考文献的注引', '（空1行）', '（空2行）', '打印删除',
        '说明:请仔细阅读', '说明：请仔细阅读', '参考文献著录规则',
        '（同[', '毕业设计论文所列', '（五号宋体：作者', '所列出的文献',
        '题名、摘要、关键词、目录', '可以在结论或讨论', '（本页为独立',
        '（可作为正文', '（小4号宋体，1.5倍行距）', '（小四号宋体，行距为固定值22磅）',
        '（每个项目开头空2字符）', '（本封面的填写内容', '重要提示',
        '餐饮', '品牌', '连锁', '加强品牌', '积极推行', '物流供应链',
    ]
    for p_elem in list(body.iter(qn('w:p'))):
        text = ''.join(p_elem.itertext()).strip()
        if not text:
            continue
        if any(m in text for m in RESIDUE_MARKERS):
            # 关键：模板有些"说明段落"里带着分页符（控制板块之间换页）。
            # 不能直接整段删除，否则会连分页符一起删掉，导致该换页的地方不换页。
            has_page_break = any(
                br.get(qn('w:type')) == 'page'
                for br in p_elem.iter(qn('w:br'))
            )
            if has_page_break:
                # 只清文字、保留分页符
                for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
                    for e in p_elem.findall(qn(tag)):
                        p_elem.remove(e)
                br_run = OxmlElement('w:r')
                br = OxmlElement('w:br')
                br.set(qn('w:type'), 'page')
                br_run.append(br)
                p_elem.append(br_run)
            else:
                parent = p_elem.getparent()
                if parent is not None:
                    parent.remove(p_elem)

    # ============ 清理连续空段落，避免产生多余空白页 ============
    # 模板为排版保留了不少空段落；清理模板示例文字后可能留下连续空段，
    # 在 Word 中表现为"莫名空白页"。这里把连续空段压缩为最多 2 个。
    # 注意：含分页符(w:br type=page)的段落不算纯空段，必须保留。
    def _is_blank(p_elem):
        if ''.join(p_elem.itertext()).strip():
            return False
        for br in p_elem.iter(qn('w:br')):
            if br.get(qn('w:type')) == 'page':
                return False
        return True

    blank_run = 0
    for p_elem in list(body.iter(qn('w:p'))):
        if p_elem.getparent() is None:
            continue
        if _is_blank(p_elem):
            blank_run += 1
            if blank_run > 1:
                parent = p_elem.getparent()
                if parent is not None:
                    parent.remove(p_elem)
        else:
            blank_run = 0

    # ============ 修复页眉：删除页眉中的空段落，避免页眉被撑成两行 ============
    # 模板 header2.xml 含两个段落：段1 为页眉文字，段2 为空段落。
    # 该空段落会使页眉区域多出一行，导致观感上"页眉换到第二行"。
    try:
        for section in doc.sections:
            for hp in [section.header, section.first_page_header, section.even_page_header]:
                if hp is None:
                    continue
                paras = list(hp.paragraphs)
                # 仅当存在有文字的页眉段落时才清理空段落（无文字的页眉如封面页眉保持原样）
                if any(pp.text.strip() for pp in paras):
                    for pp in paras:
                        if not pp.text.strip():
                            pe = pp._p
                            if pe.getparent() is not None:
                                pe.getparent().remove(pe)
    except Exception:
        pass

    # ============ 修复页眉换行：把空格串替换为"右对齐制表符" ============
    # 模板页眉用一大串三号(18pt)空格把页码推到右侧；当页码为两位数(如 27/32)时，
    # 总宽超出页眉宽度，最后一个"页"字被挤到第二行。
    # 改为：设置右对齐制表位 + 用制表符分隔，页码始终右对齐且不换行。
    try:
        for section in doc.sections:
            for hp in [section.header, section.first_page_header, section.even_page_header]:
                if hp is None:
                    continue
                for pp in hp.paragraphs:
                    if '炎黄职业技术学院毕业论文' not in pp.text:
                        continue
                    pe = pp._p
                    ppr = pe.find(qn('w:pPr'))
                    if ppr is None:
                        ppr = OxmlElement('w:pPr')
                        pe.insert(0, ppr)
                    tabs = ppr.find(qn('w:tabs'))
                    if tabs is None:
                        tabs = OxmlElement('w:tabs')
                        ppr.insert(0, tabs)
                    for e in list(tabs):
                        tabs.remove(e)
                    tb = OxmlElement('w:tab')
                    tb.set(qn('w:val'), 'right')
                    tb.set(qn('w:pos'), '8306')  # 页眉可用宽度（twips）
                    tabs.append(tb)
                    # 把"纯空格 run"删除，并在首个位置换成 tab run
                    first = True
                    for r in list(pe.findall(qn('w:r'))):
                        ts = r.findall(qn('w:t'))
                        if ts and ''.join((t.text or '') for t in ts).strip() == '':
                            if first:
                                for t in ts:
                                    r.remove(t)
                                r.append(OxmlElement('w:tab'))
                                first = False
                            else:
                                pe.remove(r)
    except Exception:
        pass

    # ============ 保存 ============
    # 全文字体统一黑色（模板要求：颜色统一设置成黑色）
    for col in body.iter(qn('w:color')):
        col.set(qn('w:val'), '000000')
    doc.save(src_path)


# ---------- 表格 cell 辅助 ----------

def clear_cell(cell):
    for p in list(cell.findall(qn('w:p'))):
        cell.remove(p)


def add_cell_para(cell, text, bold=False, font='宋体', size='21'):
    """往表格单元格添加一个段落。默认五号宋体（表格通用）；
       封面填写内容按模板要求用三号楷体_GB2312（font='楷体_GB2312', size='32'）。"""
    p = OxmlElement('w:p')
    ppr = OxmlElement('w:pPr')
    rpr_p = OxmlElement('w:rPr')
    rfonts = OxmlElement('w:rFonts')
    rfonts.set(qn('w:eastAsia'), font)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)
    rpr_p.append(rfonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), size)
    rpr_p.append(sz)
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), size)
    rpr_p.append(szcs)
    col = OxmlElement('w:color')
    col.set(qn('w:val'), '000000')
    rpr_p.append(col)
    if bold:
        b = OxmlElement('w:b')
        rpr_p.append(b)
    ppr.append(rpr_p)
    p.append(ppr)
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rfonts = OxmlElement('w:rFonts')
    rfonts.set(qn('w:eastAsia'), font)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)
    rpr.append(rfonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), size)
    rpr.append(sz)
    szcs = OxmlElement('w:szCs')
    szcs.set(qn('w:val'), size)
    rpr.append(szcs)
    col = OxmlElement('w:color')
    col.set(qn('w:val'), '000000')
    rpr.append(col)
    if bold:
        rpr.append(OxmlElement('w:b'))
    r.append(rpr)
    t = OxmlElement('w:t')
    t.text = text
    if text.startswith(' ') or text.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    p.append(r)
    cell.append(p)


def set_p_runs_text(p_elem, text):
    """把段落所有 run 的文本合并为 text"""
    for r in p_elem.findall(qn('w:r')):
        p_elem.remove(r)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = text
    r.append(t)
    p_elem.append(r)


# ---------- 长段落切分 ----------

def split_long_para(text, max_chars=500):
    """将过长段落按句号切分"""
    if len(text) <= max_chars:
        return [text]
    parts = re.split(r'(?<=。)(?=[\u4e00-\u9fff])', text)
    out = []
    cur = ''
    for p in parts:
        if len(cur) + len(p) > max_chars and cur:
            out.append(cur)
            cur = p
        else:
            cur += p
    if cur:
        out.append(cur)
    return out


# ---------- 图片段落（用 python-docx 高级插入） ----------

def make_p_picture_paragraph(src_p, img_path, width_cm=13.0):
    """克隆 src_p，去掉原 run，把图片加进来。图片用 docx 原生 add_picture 更稳。"""
    new_p = copy.deepcopy(src_p)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in new_p.findall(qn(tag)):
            new_p.remove(e)
    force_black(new_p)
    ppr = new_p.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        new_p.insert(0, ppr)
    jc = ppr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        ppr.append(jc)
    jc.set(qn('w:val'), 'center')
    # python-docx add_picture
    from docx.text.paragraph import Paragraph
    # 临时把 new_p 包装为 Paragraph 需要 parent；绕路：用 Document().add_picture 后迁移 run
    tmp_doc = Document()
    tmp_para = tmp_doc.add_paragraph()
    tmp_para.add_run().add_picture(img_path, width=Cm(width_cm))
    new_r = copy.deepcopy(tmp_para._p.findall(qn('w:r'))[0])
    new_p.append(new_r)
    return new_p


def make_caption_paragraph(src_p, caption_text):
    """图题/表题段落：克隆参考文献条目母段落（首行缩进573 固定值22磅），
       居中显示（模板：「图3-5 图与表应有相应的名称」）。"""
    new_p = copy.deepcopy(src_p)
    for tag in ('w:r', 'w:hyperlink', 'w:fldSimple'):
        for e in new_p.findall(qn(tag)):
            new_p.remove(e)
    force_black(new_p)
    ppr = new_p.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        new_p.insert(0, ppr)
    jc = ppr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        ppr.append(jc)
    jc.set(qn('w:val'), 'center')
    # 添加图题文字 run
    rpr = get_rpr(src_p)
    blacken_rpr(rpr)
    r = OxmlElement('w:r')
    if rpr is not None:
        r.append(rpr)
    t = OxmlElement('w:t')
    t.text = caption_text
    if caption_text.startswith(' ') or caption_text.endswith(' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.append(t)
    new_p.append(r)
    force_black(new_p)
    return new_p


def build_eval_table_docx(key, title):
    """构造内置 Word 表格（评价指标体系表）"""
    rows_data = [
        ('一级指标', '二级指标', '评价内容要点'),
        ('技术性能', '强度性能', '立方体抗压强度、轴心抗压强度是否达到设计强度等级'),
        ('技术性能', '变形性能', '弹性模量、收缩和徐变是否符合规范限值'),
        ('技术性能', '耐久性能', '抗冻性、抗碳化性、抗氯离子渗透性能'),
        ('技术性能', '结构性能', '构件承载力、连接节点可靠性'),
        ('施工适应性', '拌合物工作性', '坍落度经时损失、和易性、稳定性'),
        ('施工适应性', '泵送与浇筑', '泵送阻力、浇筑密实性'),
        ('施工适应性', '养护要求', '蒸汽养护制度、湿养护时间'),
        ('施工适应性', '构件生产效率', '早期强度、模具周转率'),
        ('经济与环境效益', '材料成本', '骨料成本、加工成本、运输成本'),
        ('经济与环境效益', '固废消纳量', '单位混凝土消纳建筑固废量'),
        ('经济与环境效益', '碳排放降低量', '替代天然骨料带来的隐含碳减排'),
        ('经济与环境效益', '综合经济性', '考虑补贴与外部性的综合经济性'),
    ]
    # 按论文题目定制表格内容，保证"表题"与"表格内容"一致
    if 'BIM' in title:
        rows_data = [
            ('影响维度', '作用环节', '质量提升表现'),
            ('设计协同', '统一模型下的多专业实时协同', '减少信息滞后导致的专业冲突'),
            ('构件深化', '参数化构件库与精细化节点建模', '降低构件尺寸错误与预埋件位置偏差'),
            ('管线综合', '碰撞检查与净空优化', '减少现场开洞、返工与净高不足问题'),
            ('设计变更', '参数化变更自动传播至相关视图', '降低图纸不一致性与设计变更频次'),
            ('数据贯通', '模型数据直传构件生产管理系统', '减少二次录入错误，提高构件成品率'),
        ]
    elif ('市政' in title) or ('质量管理' in title):
        rows_data = [
            ('问题类别', '主要表现', '成因归类'),
            ('构件进场', '尺寸偏差超限、表面裂缝、预埋件偏移', '材料、人员'),
            ('吊装就位', '定位偏差超限、相邻节段累计偏差', '人员、方法'),
            ('接缝防水', '卷材搭接不足、粘结不牢、收头不严', '方法、人员'),
            ('套筒连接', '排气不畅、灌浆不饱满、配合比偏差', '方法、机械'),
            ('穿插施工', '工序交叉污染、成品破坏、协调不畅', '环境、人员'),
        ]

    n_rows = len(rows_data)
    n_cols = 3
    tbl = OxmlElement('w:tbl')
    # tblPr
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)
    tblLook = OxmlElement('w:tblLook')
    tblLook.set(qn('w:val'), '04A0')
    tblPr.append(tblLook)
    tbl.append(tblPr)
    # tblGrid
    grid = OxmlElement('w:tblGrid')
    for w in (1500, 2400, 5500):
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        grid.append(gc)
    tbl.append(grid)
    for i, row in enumerate(rows_data):
        tr = OxmlElement('w:tr')
        for j, val in enumerate(row):
            tc = OxmlElement('w:tc')
            tcPr = OxmlElement('w:tcPr')
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str([1500, 2400, 5500][j]))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)
            tc.append(tcPr)
            p = OxmlElement('w:p')
            ppr = OxmlElement('w:pPr')
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center' if i == 0 or j != 2 else 'left')
            ppr.append(jc)
            p.append(ppr)
            r = OxmlElement('w:r')
            rpr = OxmlElement('w:rPr')
            rfonts = OxmlElement('w:rFonts')
            rfonts.set(qn('w:eastAsia'), '宋体')
            rfonts.set(qn('w:ascii'), 'Times New Roman')
            rfonts.set(qn('w:hAnsi'), 'Times New Roman')
            rpr.append(rfonts)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '21')
            rpr.append(sz)
            if i == 0:
                rpr.append(OxmlElement('w:b'))
            r.append(rpr)
            t = OxmlElement('w:t')
            t.text = val
            r.append(t)
            p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl