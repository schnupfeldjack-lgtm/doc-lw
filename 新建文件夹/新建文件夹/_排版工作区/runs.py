# -*- coding: utf-8 -*-
"""Inspect run-level fonts on key paragraphs."""
import sys, re
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

def rf(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None: return {}
    d = {}
    f = rpr.find(qn('w:rFonts'))
    if f is not None: d['f'] = {k.split(':')[1]: v for k, v in f.attrib.items()}
    sz = rpr.find(qn('w:sz'))
    if sz is not None: d['sz'] = int(sz.get(qn('w:val')))/2
    if rpr.find(qn('w:b')) is not None: d['b'] = 1
    c = rpr.find(qn('w:color'))
    if c is not None: d['c'] = c.get(qn('w:val'))
    va = rpr.find(qn('w:vertAlign'))
    if va is not None: d['va'] = va.get(qn('w:val'))
    return d

def main(path):
    doc = Document(path)
    body = doc.element.body
    idx = 0
    sup = 0
    for child in body.iterchildren():
        if child.tag != qn('w:p'): continue
        p = Paragraph(child, doc)
        txt = p.text.strip()
        idx += 1
        interesting = False
        if re.match(r'^\d+\s{1,2}\S', txt) and len(txt) < 40: interesting = True
        if re.match(r'^\d+[．.]\d+\s', txt) and len(txt) < 40: interesting = True
        if re.match(r'^\d+\.\d+\.\d+\s', txt) and len(txt) < 40: interesting = True
        if txt.startswith(('摘 要', '关键词', 'Abstract', 'Key words', '目  录', '结  论', '参 考 文 献', '致  谢')): interesting = True
        if re.match(r'^(图|表)\d+', txt): interesting = True
        if re.match(r'^\[\d+\]', txt): interesting = True
        if '<w:drawing>' in child.xml or '<w:pict>' in child.xml: interesting = True
        if interesting:
            runs = child.findall(qn('w:r'))
            info = [( (r.find(qn('w:t')).text or '')[:12] if r.find(qn('w:t')) is not None else '<obj>', rf(r)) for r in runs[:4]]
            print('[%d] %-46r' % (idx, txt[:44]), info)
        # count superscripts
        for r in child.findall(qn('w:r')):
            d = rf(r)
            if d.get('va') == 'superscript': sup += 1
    print('superscript runs in body:', sup)
    # table fonts
    ti = 0
    for child in body.iterchildren():
        if child.tag != qn('w:tbl'): continue
        tbl = Table(child, doc)
        fs = {}
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p._p.findall(qn('w:r')):
                        d = rf(r)
                        key = (str(d.get('f')), d.get('sz'), d.get('b'))
                        fs[key] = fs.get(key, 0) + 1
        print(f'TABLE{ti}: rows={len(tbl.rows)} runfont styles={fs}')
        ti += 1

if __name__ == '__main__':
    main(sys.argv[1])
