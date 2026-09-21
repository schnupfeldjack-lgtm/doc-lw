# -*- coding: utf-8 -*-
"""Dump DOCX structure & formatting for template analysis."""
import sys, json
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def run_font(rpr):
    if rpr is None: return {}
    rfonts = rpr.find(qn('w:rFonts'))
    d = {}
    if rfonts is not None:
        for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs', 'w:asciiTheme', 'w:eastAsiaTheme'):
            v = rfonts.get(a)
            if v: d[a.split(':')[1]] = v
    sz = rpr.find(qn('w:sz')); szcs = rpr.find(qn('w:szCs'))
    if sz is not None: d['sz'] = int(sz.get(qn('w:val')))/2.0
    if szcs is not None: d['szCs'] = int(szcs.get(qn('w:val')))/2.0
    for tag in ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps'):
        e = rpr.find(qn(tag))
        if e is not None: d[tag.split(':')[1]] = e.get(qn('w:val')) or '1'
    col = rpr.find(qn('w:color'))
    if col is not None: d['color'] = col.get(qn('w:val'))
    return d

def para_fmt(p):
    pf = p.paragraph_format
    d = {}
    try:
        d['align'] = str(pf.alignment)
        d['lineSpacing'] = pf.line_spacing
        d['lineSpacingRule'] = str(pf.line_spacing_rule)
        d['before'] = pf.space_before
        d['after'] = pf.space_after
        d['firstLine'] = pf.first_line_indent
        d['left'] = pf.left_indent
        d['right'] = pf.right_indent
    except Exception as e:
        d['err'] = str(e)
    ppr = p._p.pPr
    if ppr is not None:
        ol = ppr.find(qn('w:outlineLvl'))
        if ol is not None: d['outlineLvl'] = ol.get(qn('w:val'))
        pbdr = ppr.find(qn('w:pBdr'))
        if pbdr is not None: d['border'] = 'yes'
        shd = ppr.find(qn('w:shd'))
        if shd is not None: d['shd'] = shd.get(qn('w:fill'))
        snap = ppr.find(qn('w:snapToGrid'))
        if snap is not None: d['snapToGrid'] = snap.get(qn('w:val'))
        pb = ppr.find(qn('w:pageBreakBefore'))
        if pb is not None: d['pageBreakBefore'] = pb.get(qn('w:val'))
        ka = ppr.find(qn('w:keepNext'))
        if ka is not None: d['keepNext'] = ka.get(qn('w:val'))
        tabs = ppr.find(qn('w:tabs'))
        if tabs is not None:
            d['tabs'] = [(t.get(qn('w:val')), t.get(qn('w:pos')), t.get(qn('w:leader'))) for t in tabs]
        ind = ppr.find(qn('w:ind'))
        if ind is not None:
            d['ind_attr'] = {k.split(':')[1]: v for k, v in ind.attrib.items()}
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None: d['pPr_rPr'] = run_font(rpr)
        numpr = ppr.find(qn('w:numPr'))
        if numpr is not None:
            ilvl = numpr.find(qn('w:ilvl')); numid = numpr.find(qn('w:numId'))
            d['numPr'] = (ilvl.get(qn('w:val')) if ilvl is not None else None,
                          numid.get(qn('w:val')) if numid is not None else None)
        sectpr = ppr.find(qn('w:sectPr'))
        if sectpr is not None: d['sectPr_in_pPr'] = True
    return d

def runs_info(p):
    out = []
    for r in p.runs:
        t = r.text
        info = {'t': t[:80], 'len': len(t), 'font': run_font(r._r.rPr)}
        # detect objects
        xml = r._r.xml
        if 'w:object' in xml or 'm:oMath' in xml or '<m:oMathPara' in xml:
            info['obj'] = 'omml' if 'm:oMath' in xml else 'object'
        if '<w:drawing>' in xml or '<w:pict>' in xml: info['img'] = True
        if 'w:fldChar' in xml: info['field'] = True
        if 'w:instrText' in xml:
            import re
            info['instr'] = re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', xml)
        out.append(info)
    return out

def sect_info(doc):
    out = []
    for i, s in enumerate(doc.sections):
        d = {'idx': i, 'start_type': str(s.start_type), 'page_w': s.page_width, 'page_h': s.page_height,
             'orient': str(s.orientation),
             'left': s.left_margin, 'right': s.right_margin, 'top': s.top_margin, 'bottom': s.bottom_margin,
             'gutter': s.gutter, 'header_dist': s.header_distance, 'footer_dist': s.footer_distance,
             'diff_first': s.different_first_page_header_footer,
             }
        try:
            d['odd_even'] = s.odd_and_even_pages_header_footer
        except Exception:
            d['odd_even'] = None
        sp = s._sectPr
        pg = sp.find(qn('w:pgNumType'))
        if pg is not None:
            d['pgNumType'] = {k.split(':')[1]: v for k, v in pg.attrib.items()}
        cols = sp.find(qn('w:cols'))
        if cols is not None: d['cols'] = cols.get(qn('w:num'))
        docgrid = sp.find(qn('w:docGrid'))
        if docgrid is not None:
            d['docGrid'] = {k.split(':')[1]: v for k, v in docgrid.attrib.items()}
        # headers footers
        for hf in ('header', 'footer'):
            for t in ('default', 'first', 'even'):
                e = getattr(s, f'{t}_{hf}', None)
                if e is not None:
                    txt = []
                    for p in e.paragraphs:
                        if p.text.strip(): txt.append(p.text.strip()[:60])
                    if txt: d[f'{t}_{hf}'] = txt
        lnk = [e is not None and e.is_linked_to_previous for e in (s.header, s.footer)]
        d['hf_linked'] = lnk
        out.append(d)
    return out

def body_text(doc):
    """Iterate body-level items in order: paragraphs and tables."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    items = []
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            items.append(('p', Paragraph(child, doc)))
        elif child.tag == qn('w:tbl'):
            items.append(('t', Table(child, doc)))
        elif child.tag == qn('w:sectPr'):
            items.append(('sect', None))
    return items

def main(path, maxpar=None, only=None):
    doc = Document(path)
    print('=== FILE:', path)
    print('=== SECTIONS ===')
    for s in sect_info(doc):
        print(json.dumps(s, ensure_ascii=False, default=str))
    print('=== STYLES ===')
    try:
        styles = doc.styles
        for st in styles:
            if st.type is not None and str(st.type) in ('PARAGRAPH (1)', 'TABLE (5)'):
                f = st.font
                pf = st.paragraph_format
                sd = {'name': st.name, 'sz': f.size, 'bold': f.bold,
                      'ascii': None, 'ea': None}
                try:
                    rpr = st.element.find(qn('w:rPr'))
                    if rpr is not None:
                        sd['rpr'] = run_font(rpr)
                    ppr = st.element.find(qn('w:pPr'))
                    if ppr is not None:
                        jd = {}
                        for tag in ('w:spacing', 'w:ind', 'w:jc'):
                            e = ppr.find(qn(tag))
                            if e is not None: jd[tag.split(':')[1]] = {k.split(':')[1]: v for k, v in e.attrib.items()}
                        ol = ppr.find(qn('w:outlineLvl'))
                        if ol is not None: jd['outlineLvl'] = ol.get(qn('w:val'))
                        if jd: sd['ppr'] = jd
                except Exception as e:
                    sd['err'] = str(e)
                print(json.dumps(sd, ensure_ascii=False, default=str))
    except Exception as e:
        print('styles err', e)
    print('=== BODY ITEMS ===')
    items = body_text(doc)
    for i, (kind, obj) in enumerate(items):
        if kind == 'sect':
            print(f'[{i}] <sectPr>')
            continue
        if kind == 't':
            print(f'[{i}] TABLE rows={len(obj.rows)} cols={len(obj.columns)} first_cell="{obj.rows[0].cells[0].text[:40] if obj.rows else ""}"')
            continue
        p = obj
        txt = p.text
        d = {'i': i, 'style': p.style.name if p.style else None, 'text': txt[:100], 'len': len(txt),
             'fmt': para_fmt(p), 'runs': runs_info(p)}
        print(json.dumps(d, ensure_ascii=False, default=str))
        if maxpar and i > maxpar: break

if __name__ == '__main__':
    main(sys.argv[1], maxpar=int(sys.argv[2]) if len(sys.argv) > 2 else None)
