import zipfile, re, collections
from lxml import etree
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
files = [BASE + r"\炎黄职业技术学院毕业论文模板(1).docx",
 BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
 BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
 BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"]
for f in files:
    z = zipfile.ZipFile(f)
    xml = z.read('word/document.xml')
    root = etree.fromstring(xml)
    ids = root.xpath('//@w14:paraId', namespaces={'w14':'http://schemas.microsoft.com/office/word/2010/wordml'})
    tids = root.xpath('//@w14:textId', namespaces={'w14':'http://schemas.microsoft.com/office/word/2010/wordml'})
    c = collections.Counter(ids)
    dup = {k:v for k,v in c.items() if v>1}
    ct = collections.Counter(tids)
    dupt = {k:v for k,v in ct.items() if v>1}
    print(f"{f[-30:]}: paraId总数={len(ids)} 重复值={len(dup)} 重复实例={sum(dup.values())-len(dup)} | textId重复={len(dupt)}")
    # 列出前几个重复
    for k,v in list(dup.items())[:5]:
        print("    重复 paraId", k, v, "次")
