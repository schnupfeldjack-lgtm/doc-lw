# -*- coding: utf-8 -*-
from pathlib import Path
import re
import shutil
import tempfile
import zipfile
from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
DIR = ROOT / "新建文件夹/新建文件夹/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例"
BASE = DIR / "BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"
OUT = DIR / "BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例_图片修复版.docx"
FIG3 = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p3/fig3_1.png"
FIG4 = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p3/fig4_1.png"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_EMBED = "{%s}embed" % NS["r"]

def text_of(p):
    return "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()

def norm(s):
    return re.sub(r"\s+", "", s).replace("–", "-").replace("—", "-")

def find_caption(root, prefix):
    for p in root.xpath("//w:body//w:p", namespaces=NS):
        if norm(text_of(p)).startswith(prefix):
            return p
    raise RuntimeError("未找到图题：" + prefix)

def prev_drawing(caption):
    p = caption.getprevious()
    for _ in range(8):
        if p is None:
            break
        if p.tag == "{%s}p" % NS["w"] and p.xpath(".//w:drawing", namespaces=NS):
            return p
        p = p.getprevious()
    raise RuntimeError("图题前未找到图片段落：" + text_of(caption))

def next_rids(rels, count=2):
    nums = []
    for r in rels:
        m = re.fullmatch(r"rId(\d+)", r.get("Id", ""))
        if m:
            nums.append(int(m.group(1)))
    n = max(nums or [0]) + 1
    return [f"rId{n+i}" for i in range(count)]

def add_rel(rels, rid, target):
    el = etree.Element("{%s}Relationship" % NS["rel"])
    el.set("Id", rid)
    el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    el.set("Target", target)
    rels.append(el)

def main():
    for p in (BASE, FIG3, FIG4):
        if not p.exists():
            raise FileNotFoundError(p)

    with tempfile.TemporaryDirectory() as td:
        work = Path(td)
        with zipfile.ZipFile(BASE) as z:
            z.extractall(work)

        doc_path = work / "word/document.xml"
        rel_path = work / "word/_rels/document.xml.rels"
        media = work / "word/media"
        media.mkdir(parents=True, exist_ok=True)

        doc = etree.parse(str(doc_path))
        rels_doc = etree.parse(str(rel_path))
        rels = rels_doc.getroot()
        rid3, rid4 = next_rids(rels, 2)

        cap3 = find_caption(doc.getroot(), "图3-1BIM协同设计实施方案主要环节")
        cap4 = find_caption(doc.getroot(), "图4-1BIM协同设计对结构设计质量的影响机理")
        pic3 = prev_drawing(cap3)
        pic4 = prev_drawing(cap4)

        # 原文件两张图错误地都指向 rId9（fontTable），因此 Word/WPS 显示空白。
        for blip in pic3.xpath(".//a:blip", namespaces=NS):
            blip.set(R_EMBED, rid3)
        for blip in pic4.xpath(".//a:blip", namespaces=NS):
            blip.set(R_EMBED, rid4)

        # 保留原图框尺寸和位置，只替换图片数据。
        shutil.copy2(FIG3, media / "fig3_1_repaired.png")
        shutil.copy2(FIG4, media / "fig4_1_repaired.png")
        add_rel(rels, rid3, "media/fig3_1_repaired.png")
        add_rel(rels, rid4, "media/fig4_1_repaired.png")

        doc.write(str(doc_path), encoding="UTF-8", xml_declaration=True, standalone="yes")
        rels_doc.write(str(rel_path), encoding="UTF-8", xml_declaration=True, standalone="yes")

        with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
            for p in work.rglob("*"):
                if p.is_file():
                    z.write(p, p.relative_to(work).as_posix())

    # 严格校验：必须真实嵌入两张 PNG，且两个图片关系分别存在。
    with zipfile.ZipFile(OUT) as z:
        names = set(z.namelist())
        assert "word/media/fig3_1_repaired.png" in names
        assert "word/media/fig4_1_repaired.png" in names
        dxml = z.read("word/document.xml").decode("utf-8")
        rxml = z.read("word/_rels/document.xml.rels").decode("utf-8")
        assert f'r:embed="{rid3}"' in dxml
        assert f'r:embed="{rid4}"' in dxml
        assert f'Id="{rid3}"' in rxml and 'fig3_1_repaired.png' in rxml
        assert f'Id="{rid4}"' in rxml and 'fig4_1_repaired.png' in rxml
        assert len(z.read("word/media/fig3_1_repaired.png")) > 100000
        assert len(z.read("word/media/fig4_1_repaired.png")) > 100000

    print("OK", OUT, OUT.stat().st_size, rid3, rid4)

if __name__ == "__main__":
    main()
