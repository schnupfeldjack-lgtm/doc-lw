# -*- coding: utf-8 -*-
"""构造只含指定 body 子元素的临时 docx，并用 Word 测试能否打开
用法: python t_build.py <源docx> <索引列表,逗号分隔> [--rows N]
"""
import zipfile, copy, sys, re
from lxml import etree
import win32com.client as win32

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

SRC = sys.argv[1]
IDX = [int(x) for x in sys.argv[2].split(',') if x.strip() != '']
ROWS = None
if '--rows' in sys.argv:
    ROWS = int(sys.argv[sys.argv.index('--rows') + 1])

TMP = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\t_build_tmp.docx"

z = zipfile.ZipFile(SRC)
items = z.infolist()
parts = {it.filename: z.read(it.filename) for it in items}
z.close()
root = etree.fromstring(parts['word/document.xml'])
body = root.find(q('body'))
children = list(body)

newbody = etree.Element(q('body'))
for i in IDX:
    el = copy.deepcopy(children[i])
    if ROWS is not None and el.tag == q('tbl'):
        for tr in el.findall(q('tr'))[ROWS:]:
            el.remove(tr)
    newbody.append(el)
last = children[-1]
if last.tag == q('sectPr') or (last.find(q('pPr')) is not None and last.find(q('pPr')).find(q('sectPr')) is not None):
    newbody.append(copy.deepcopy(last))
newroot = copy.deepcopy(root)
newroot.replace(newroot.find(q('body')), newbody)
data = etree.tostring(newroot, xml_declaration=True, encoding='UTF-8', standalone=True)
o = zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED)
for it in items:
    o.writestr(it, data if it.filename == 'word/document.xml' else parts[it.filename])
o.close()

w = win32.DispatchEx("Word.Application"); w.Visible = False; w.DisplayAlerts = 0
try:
    d = w.Documents.Open(TMP, ReadOnly=True)
    print("OK  索引", IDX, "rows=", ROWS)
    d.Close(False)
except Exception as e:
    print("FAIL 索引", IDX, "rows=", ROWS, "|", str(e)[:60])
w.Quit()
