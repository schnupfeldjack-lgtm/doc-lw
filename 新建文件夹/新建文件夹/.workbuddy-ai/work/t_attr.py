# -*- coding: utf-8 -*-
"""查找 w14:paraId / textId 出现在非 w:p 元素上的情况"""
import zipfile, sys
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
def local(t):
    return t.split('}')[-1]
for f in sys.argv[1:]:
    z = zipfile.ZipFile(f)
    root = etree.fromstring(z.read('word/document.xml'))
    bad = {}
    for el in root.iter():
        for a in (f'{{{W14}}}paraId', f'{{{W14}}}textId'):
            if a in el.attrib:
                if local(el.tag) != 'p':
                    key = (local(el.tag), a.split('}')[-1])
                    bad[key] = bad.get(key, 0) + 1
    print(f"{f[-26:]}: 非段落元素上的 paraId/textId -> {bad if bad else '无'}")
