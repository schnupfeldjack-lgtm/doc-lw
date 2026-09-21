# -*- coding: utf-8 -*-
"""
修复导致 Word 报"文件已损坏"的结构问题：
1) 每个 <w:tc> 必须至少含一个 <w:p>（之前清理空段落时把表格单元格内的空段也删了）
2) 每个 <w:tr> 至少含一个 <w:tc>；每个 <w:tbl> 至少含一个 <w:tr>
直接在 zip 层面重写 word/document.xml（保留其余部件不变，避免 python-docx 二次改动）
"""
import zipfile, copy, os, sys
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

def fix(path):
    z = zipfile.ZipFile(path)
    items = z.infolist()
    parts = {it.filename: z.read(it.filename) for it in items}
    z.close()
    root = etree.fromstring(parts['word/document.xml'])
    n_cell = n_row = 0
    for tbl in root.iter(q('tbl')):
        # 取一个模板段落用于补空单元格
        sample = None
        for p in tbl.iter(q('p')):
            sample = p
            break
        trs = tbl.findall(q('tr'))
        if not trs:
            continue
        for tr in trs:
            tcs = tr.findall(q('tc'))
            if not tcs:
                tc = etree.SubElement(tr, q('tc'))
                tc.append(etree.Element(q('tcPr')))
                tcs = [tc]
                n_row += 1
            for tc in tcs:
                if tc.find(q('p')) is None:
                    if sample is not None:
                        np = copy.deepcopy(sample)
                        # 清空所有 run，只留 pPr
                        for tag in ('w:r', 'w:hyperlink', 'w:fldSimple', 'w:bookmarkStart',
                                    'w:bookmarkEnd', 'w:proofErr'):
                            for e in np.findall(q(tag[2:])):
                                np.remove(e)
                    else:
                        np = etree.Element(q('p'))
                        np.append(etree.Element(q('pPr')))
                    tc.append(np)
                    n_cell += 1
    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    o = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        o.writestr(it, data if it.filename == 'word/document.xml' else parts[it.filename])
    o.close()
    return n_cell, n_row

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
    BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
    BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx",
]
for p in DOCS:
    c, r = fix(p)
    print(f"{os.path.basename(p)[:26]}: 补空单元格 {c} 个，补空行 {r} 个")
print("完成")
