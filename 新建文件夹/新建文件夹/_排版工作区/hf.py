# -*- coding: utf-8 -*-
"""Extract headers/footers (incl. tables & field codes) and cover tables."""
import sys, re
from docx import Document
from docx.oxml.ns import qn

def elem_text(e):
    """concat all w:t in element, mark fields and tabs"""
    out = []
    for node in e.iter():
        tag = node.tag
        if tag == qn('w:t'):
            out.append(node.text or '')
        elif tag == qn('w:tab'):
            out.append('\\t')
        elif tag == qn('w:instrText'):
            out.append('{FIELD:' + (node.text or '').strip() + '}')
        elif tag == qn('w:br'):
            out.append('\\n')
    return ''.join(out)

def para_desc(p):
    ppr = p.find(qn('w:pPr'))
    info = ''
    if ppr is not None:
        jc = ppr.find(qn('w:jc'))
        tabs = ppr.find(qn('w:tabs'))
        rpr = ppr.find(qn('w:rPr'))
        bits = []
        if jc is not None: bits.append('jc=' + jc.get(qn('w:val')))
        if tabs is not None:
            bits.append('tabs=' + str([(t.get(qn('w:val')), t.get(qn('w:pos')), t.get(qn('w:leader'))) for t in tabs]))
        if rpr is not None:
            rf = rpr.find(qn('w:rFonts')); sz = rpr.find(qn('w:sz')); b = rpr.find(qn('w:b'))
            fbits = []
            if rf is not None: fbits.append('fonts=' + str({k.split(':')[1]: v for k, v in rf.attrib.items()}))
            if sz is not None: fbits.append('sz=' + str(int(sz.get(qn('w:val')))/2))
            if b is not None: fbits.append('b=1')
            bits.append(' '.join(fbits))
        info = ' | '.join(bits)
    return elem_text(p), info

def dump_hf(doc):
    for si, sect in enumerate(doc.sections):
        for kind in ('header', 'footer'):
            for ref in sect._sectPr.findall(qn(f'w:{kind}Reference')):
                rid = ref.get(qn('r:id'))
                typ = ref.get(qn('w:type'))
                try:
                    part = sect.part.related_parts[rid]
                    el = part.element
                except Exception as e:
                    print(f'--- sect{si} {kind} type={typ} ERR {e}')
                    continue
                t = typ or 'default'
                texts = []
                for p in el.findall(qn('w:p')):
                    t, info = para_desc(p)
                    if t.strip() or info: texts.append(('P', t, info))
                for tbl in el.findall(qn('w:tbl')):
                    for tr in tbl.findall(qn('w:tr')):
                        cells = []
                        for tc in tr.findall(qn('w:tc')):
                            cells.append(elem_text(tc))
                        texts.append(('TR', ' || '.join(cells), ''))
                if texts:
                    print(f'--- sect{si} {kind}/{t}')
                    for k, t2, info in texts:
                        print(f'    {k}: {t2!r}  {info}')

def dump_tables(doc, idxs):
    body = doc.element.body
    ti = 0
    for child in body.iterchildren():
        if child.tag == qn('w:tbl'):
            if ti in idxs:
                print(f'=== TABLE #{ti}')
                for tr in child.findall(qn('w:tr')):
                    cells = []
                    for tc in tr.findall(qn('w:tc')):
                        txt = elem_text(tc)
                        tcpr = tc.find(qn('w:tcPr'))
                        w = tcpr.find(qn('w:tcW')).get(qn('w:w')) if tcpr is not None and tcpr.find(qn('w:tcW')) is not None else '?'
                        vm = tcpr.find(qn('w:vAlign')).get(qn('w:val')) if tcpr is not None and tcpr.find(qn('w:vAlign')) is not None else ''
                        cells.append(f"[{w}/{vm}]{txt}")
                    print('   | ' + ' || '.join(cells))
                # table properties
                tblpr = child.find(qn('w:tblPr'))
                if tblpr is not None:
                    print('   tblPr:', {k.split('}')[-1]: v for k, v in tblpr.attrib.items()})
                    jc = tblpr.find(qn('w:jc'))
                    if jc is not None: print('   tblJc:', jc.get(qn('w:val')))
                    borders = tblpr.find(qn('w:tblBorders'))
                    if borders is not None:
                        print('   borders:', {b.tag.split('}')[-1]: {k.split('}')[-1]: v for k, v in b.attrib.items()} for b in borders})
            ti += 1

if __name__ == '__main__':
    doc = Document(sys.argv[1])
    print('###### HEADERS/FOOTERS')
    dump_hf(doc)
    idxs = [int(x) for x in sys.argv[2].split(',')] if len(sys.argv) > 2 else []
    print('###### TABLES')
    dump_tables(doc, idxs)
