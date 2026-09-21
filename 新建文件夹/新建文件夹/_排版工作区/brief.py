# -*- coding: utf-8 -*-
import json, sys
f = sys.argv[1]; a = int(sys.argv[2]); b = int(sys.argv[3])
lines = open(f, encoding='utf-8').read().split('\n')
for ln in lines[a:b]:
    if ln.startswith('{"i"'):
        d = json.loads(ln); fm = d['fmt']
        def pt(v):
            return round(v/12700.0, 1) if isinstance(v, (int, float)) else v
        print(f"[{d['i']}] {d['text'][:70]!r}")
        print(f"    style={d['style']} align={fm.get('align')} ls={fm.get('lineSpacing')}/{fm.get('lineSpacingRule')} bef={pt(fm.get('before'))} aft={pt(fm.get('after'))} fl={pt(fm.get('firstLine'))} L={pt(fm.get('left'))} R={pt(fm.get('right'))}")
        x = {k: v for k, v in fm.items() if k in ('outlineLvl', 'tabs', 'pageBreakBefore', 'keepNext', 'snapToGrid', 'ind_attr', 'numPr', 'sectPr_in_pPr')}
        if x: print('    ', x)
        print('    pPr_rPr=', fm.get('pPr_rPr'))
        r0 = d['runs'][0] if d['runs'] else {}
        if d['runs'] and len(d['runs']) > 1:
            print('    runs fonts=', [(r['t'][:14], r['font']) for r in d['runs'][:6]])
    elif ln.startswith('['):
        print(ln[:160])
