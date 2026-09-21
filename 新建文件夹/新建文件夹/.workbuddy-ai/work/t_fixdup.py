import zipfile, shutil, os, re, uuid
from lxml import etree
src = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/再生混凝土在装配式建筑中的应用评价——以住宅项目为例/再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx"
dst = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/test_nodeid.docx"
z = zipfile.ZipFile(src)
data = z.read('word/document.xml')
root = etree.fromstring(data)
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
n = 0
for el in root.iter():
    for a in (f'{{{W14}}}paraId', f'{{{W14}}}textId'):
        if a in el.attrib:
            del el.attrib[a]; n += 1
new = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
out = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    b = z.read(it.filename)
    if it.filename == 'word/document.xml':
        b = new
    out.writestr(it, b)
out.close()
print("去除属性数:", n, "->", dst)
