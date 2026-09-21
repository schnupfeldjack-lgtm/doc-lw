# -*- coding: utf-8 -*-
"""Compact structure inventory of a thesis docx."""
import sys, re
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

def pt(v):
    return round(v/12700.0, 1) if isinstance(v, (int, float)) else v

def elem_text(e):
    out = []
    for n in e.iter():
        if n.tag == qn('w:t'): out.append(n.text or '')
        elif n.tag == qn('w:tab'): out.append('\\t')
        elif n.tag == qn('w:instrText'): out.append('{' + (n.text or '').strip() + '}')
    return ''.join(out)

def run_font(rpr):
    if rpr is None: return {}
    d = {}
    rf = rpr.find(qn('w:rFonts'))
    if rf is not None: d['f'] = {k.split(':')[1]: v for k, v in rf.attrib.items() if k.split(':')[1] in ('ascii', 'hAnsi', 'eastAsia', 'hint')}
    sz = rpr.find(qn('w:sz'))
    if sz is not None: d['sz'] = int(sz.get(qn('w:val')))/2
    if rpr.find(qn('w:b')) is not None: d['b'] = 1
    c = rpr.find(qn('w:color'))
    if c is not None: d['c'] = c.get(qn('w:val'))
    return d

def pf(p):
    f = p.paragraph_format
    s = []
    try:
        s.append('align=%s' % f.alignment)
        s.append('ls=%s/%s' % (f.line_spacing, f.line_spacing_rule))
        s.append('bef=%s aft=%s' % (pt(f.space_before), pt(f.space_after)))
        s.append('fl=%s L=%s R=%s' % (pt(f.first_line_indent), pt(f.left_indent), pt(f.right_indent)))
    except Exception as e:
        s.append('err')
    ppr = p._p.pPr
    if ppr is not None:
        ol = ppr.find(qn('w:outlineLvl'))
        if ol is not None: s.append('ol=%s' % ol.get(qn('w:val')))
        if ppr.find(qn('w:pageBreakBefore')) is not None: s.append('PBBefore')
        ind = ppr.find(qn('w:ind'))
        if ind is not None: s.append('ind=%s' % {k.split(':')[1]: v for k, v in ind.attrib.items()})
        rpr = ppr.find(qn('w:rPr'))
        if rpr is not None: s.append('pPr=%s' % run_font(rpr))
    return ' '.join(x for x in s if not x.endswith('None'))

def main(path):
    doc = Document(path)
    print('### FILE', path)
    xml = doc.element.body.xml
    print('images(drawing/pict)=%d  oMath=%d  fields(instrText)=%d' % (
        xml.count('<w:drawing>') + xml.count('<w:pict>'), xml.count('<m:oMath'), xml.count('w:instrText')))
    print('### SECTIONS')
    for i, s in enumerate(doc.sections):
        print(' sect%d start=%s w=%s h=%s L=%s R=%s T=%s B=%s gut=%s hdrD=%s ftrD=%s first_diff=%s' % (
            i, s.start_type, pt(s.page_width), pt(s.page_height), pt(s.left_margin), pt(s.right_margin),
            pt(s.top_margin), pt(s.bottom_margin), pt(s.gutter), pt(s.header_distance), pt(s.footer_distance),
            s.different_first_page_header_footer))
        sp = s._sectPr
        pg = sp.find(qn('w:pgNumType'))
        if pg is not None: print('   pgNumType', {k.split(':')[1]: v for k, v in pg.attrib.items()})
        dg = sp.find(qn('w:docGrid'))
        if dg is not None: print('   docGrid', {k.split(':')[1]: v for k, v in dg.attrib.items()})
        for kind in ('header', 'footer'):
            for ref in sp.findall(qn(f'w:{kind}Reference')):
                rid = ref.get(qn('r:id')); typ = ref.get(qn('w:type'))
                try:
                    part = s.part.related_parts[rid]
                    print(f'   {kind}/{typ or "default"}:', repr(elem_text(part.element))[:200])
                except Exception as e:
                    print('   ', kind, typ, 'ERR', e)
        print('   hdr_linked=%s ftr_linked=%s' % (s.header.is_linked_to_previous, s.footer.is_linked_to_previous))
    print('### BODY')
    ti = 0
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:tbl'):
            tbl = Table(child, doc)
            rows = ' || '.join('|'.join(c.text.replace('\n', ' ')[:24] for c in r.cells) for r in tbl.rows[:2])
            print(f'[T{ti}] rows={len(tbl.rows)} cols={len(tbl.columns)} :: {rows[:160]}')
            ti += 1
        elif child.tag == qn('w:p'):
            p = Paragraph(child, doc)
            txt = p.text
            flags = []
            if 'm:oMath' in child.xml or 'm:oMathPara' in child.xml: flags.append('MATH')
            if '<w:drawing>' in child.xml or '<w:pict>' in child.xml: flags.append('IMG')
            if 'w:instrText' in child.xml: flags.append('FLD:' + ';'.join(re.findall(r'<w:instrText[^>]*>([^<]*)</w:instrText>', child.xml))[:80])
            if child.find(qn('w:bookmarkStart')) is not None: flags.append('BM')
            print('[P] %-58r | %s | %s %s' % (txt[:56], p.style.name, pf(p), ' '.join(flags)))
        elif child.tag == qn('w:sectPr'):
            print('[SECTPR]')
    # footnotes/endnotes?
    try:
        if doc.part.package.part_related_by('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'):
            print('has footnotes part')
    except Exception:
        pass

if __name__ == '__main__':
    main(sys.argv[1])
