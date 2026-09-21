import zipfile, re
from lxml import etree
R='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
files = [BASE + r"\炎黄职业技术学院毕业论文模板(1).docx",
 BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
 BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
 BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"]
RID = re.compile(r'r:(?:id|embed|link|pict|media)="([^"]+)"')
for f in files:
    z = zipfile.ZipFile(f)
    names = set(z.namelist())
    print("==", f[-26:])
    # content types
    ct = etree.fromstring(z.read('[Content_Types].xml'))
    declared = {o.get('PartName') for o in ct.findall(f'{{{CT}}}Override')}
    for n in names:
        if n.endswith('.xml') and not n.startswith('_rels') and n != '[Content_Types].xml':
            if '/' + n not in declared:
                print("   ★未在ContentTypes声明:", n)
    # rels
    for rn in [n for n in names if n.endswith('.rels')]:
        rel = etree.fromstring(z.read(rn))
        ids = {r.get('Id') for r in rel.findall(f'{{{R}}}Relationship')}
        targets = {r.get('Target') for r in rel.findall(f'{{{R}}}Relationship')}
        part = rn.replace('_rels/','').replace('.rels','')
        if part not in names and rn != '_rels/.rels':
            print("   ★rels对应部件不存在:", rn, part)
        if part in names:
            data = z.read(part).decode('utf-8', 'ignore')
            used = set(RID.findall(data))
            miss = used - ids
            if miss:
                print(f"   ★{part} 引用了不存在的 rId: {miss} (rels中共有 {len(ids)})")
        for t in targets:
            if t.startswith('/'):
                tn = t.lstrip('/')
                if tn not in names: print("   ★Target缺失:", rn, t)
            elif not t.startswith('http'):
                import posixpath
                tn = posixpath.normpath(posixpath.join(posixpath.dirname(part), t))
                if tn not in names: print("   ★Target缺失:", rn, t, '->', tn)
    print("   ok")
