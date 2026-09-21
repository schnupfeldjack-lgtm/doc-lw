# -*- coding: utf-8 -*-
import zipfile, re, sys
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"
SRC = sys.argv[1]
IDX = int(sys.argv[2])
LIM = int(sys.argv[3]) if len(sys.argv) > 3 else 6000
z = zipfile.ZipFile(SRC)
root = etree.fromstring(z.read('word/document.xml'))
body = root.find(q('body'))
ch = list(body)
el = ch[IDX]
s = etree.tostring(el, pretty_print=True).decode()
s = re.sub(r'\sxmlns:\w+="[^"]*"', '', s)
print(s[:LIM])
