# -*- coding: utf-8 -*-
"""二分定位导致 Word 报"文件已损坏"的正文元素"""
import zipfile, os, sys
from lxml import etree
import win32com.client as win32

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

SRC = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"
TMP = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹\.workbuddy-ai\work\bisect_tmp.docx"

z = zipfile.ZipFile(SRC)
items = z.infolist()
parts = {it.filename: z.read(it.filename) for it in items}
z.close()

root = etree.fromstring(parts['word/document.xml'])
body = root.find(q('body'))
children = list(body)

def build(keep_idx, out=TMP):
    """保留 body 前 keep_idx 个普通子元素 + 末尾 sectPr"""
    import copy
    newbody = etree.Element(q('body'))
    for c in children[:keep_idx]:
        newbody.append(copy.deepcopy(c))
    last = children[-1]
    if last.tag == q('sectPr'):
        newbody.append(copy.deepcopy(last))
    elif children[-1].find(q('pPr')) is not None and children[-1].find(q('pPr')).find(q('sectPr')) is not None:
        newbody.append(copy.deepcopy(last))
    newroot = copy.deepcopy(root)
    oldbody = newroot.find(q('body'))
    newroot.replace(oldbody, newbody)
    data = etree.tostring(newroot, xml_declaration=True, encoding='UTF-8', standalone=True)
    o = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        o.writestr(it, data if it.filename == 'word/document.xml' else parts[it.filename])
    o.close()

word = win32.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
def ok():
    try:
        d = word.Documents.Open(TMP, ReadOnly=True)
        d.Close(False)
        return True
    except Exception:
        return False

total = len(children)
print("body 子元素总数:", total)
lo, hi = 1, total
# 先确认全量失败
build(total)
print("全量:", "OK" if ok() else "FAIL")
# 找到第一个失败的位置
while lo < hi:
    mid = (lo + hi) // 2
    build(mid)
    r = ok()
    print(f"  前{mid}个: {'OK' if r else 'FAIL'}")
    if r: lo = mid + 1
    else: hi = mid
print("★ 第一个导致失败的元素索引(0-based):", lo - 1)
c = children[lo - 1]
s = etree.tostring(c).decode()[:1500]
import re
s = re.sub(r'\sxmlns:\w+="[^"]*"', '', s)
print("该元素XML:\n", s)
word.Quit()
