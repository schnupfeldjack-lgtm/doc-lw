# -*- coding: utf-8 -*-
"""
规范化 OOXML：
1) w:pPr / w:rPr 子元素按 ECMA-376 schema 顺序重排（乱序会让 Word 报"文件已损坏"）
2) w14:paraId / w14:textId 去重（复制母段落导致大量重复）
就地重写 docx 内的所有 xml 部件。
"""
import zipfile, os, random
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'

PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr',
    'suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap',
    'overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid',
    'spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection',
    'textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange']

RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline',
    'shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w',
    'kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl',
    'cs','em','lang','eastAsianLayout','specVanish','oMath']

def order_key(order):
    idx = {f'{{{W}}}{n}': i for i, n in enumerate(order)}
    return lambda el: idx.get(el.tag, len(order))

def sort_children(el, order):
    kids = list(el)
    if not kids:
        return 0
    keys = order_key(order)
    srt = sorted(kids, key=keys)
    if [k.tag for k in srt] == [k.tag for k in kids]:
        return 0
    for k in kids:
        el.remove(k)
    for k in srt:
        el.append(k)
    return 1

def norm_part(data):
    root = etree.fromstring(data)
    n = 0
    for el in root.iter():
        if el.tag == f'{{{W}}}pPr':
            n += sort_children(el, PPR_ORDER)
        elif el.tag == f'{{{W}}}rPr':
            n += sort_children(el, RPR_ORDER)
    # paraId / textId 去重
    seen = set()
    m = 0
    for el in root.iter():
        for a in (f'{{{W14}}}paraId', f'{{{W14}}}textId'):
            if a in el.attrib:
                v = el.attrib[a]
                if v in seen:
                    el.attrib[a] = '%08X' % random.getrandbits(32)
                    m += 1
                else:
                    seen.add(v)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), n, m

def process(path):
    z = zipfile.ZipFile(path)
    items = z.infolist()
    newdata = {}
    tot_s = tot_i = 0
    for it in items:
        b = z.read(it.filename)
        if it.filename.endswith('.xml') or it.filename.endswith('.rels'):
            try:
                nb, s, i2 = norm_part(b)
                newdata[it.filename] = nb
                tot_s += s; tot_i += i2
            except Exception:
                newdata[it.filename] = b
        else:
            newdata[it.filename] = b
    z.close()
    out = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        out.writestr(it, newdata[it.filename])
    out.close()
    return tot_s, tot_i

if __name__ == '__main__':
    import sys
    BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
    DOCS = [
        BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
        BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
        BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx",
    ]
    for p in DOCS:
        s, i = process(p)
        print(f"{os.path.basename(p)[:26]}: 重排 {s} 处，paraId 去重 {i} 处")
    print("完成")
