# -*- coding: utf-8 -*-
import os, re, shutil, tempfile, zipfile, copy
from pathlib import Path
from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / "新建文件夹/新建文件夹/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"
SOURCE = ROOT / "新建文件夹/新建文件夹/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例_按模板排版.docx"
OUT = ROOT / "新建文件夹/新建文件夹/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例/BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例_图片修复版.docx"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "v": "urn:schemas-microsoft-com:vml",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
R = "{%s}" % NS["r"]
REL = "{%s}" % NS["rel"]
CT = "{%s}" % NS["ct"]

CAPTIONS = [
    ("图3-1", re.compile(r"^图\s*3[–—-]1\s*BIM\s*协同设计实施方案主要环节")),
    ("图4-1", re.compile(r"^图\s*4[–—-]1\s*BIM\s*协同设计对结构设计质量的影响机理")),
]

def unzip(src, dst):
    with zipfile.ZipFile(src) as z:
        z.extractall(dst)

def repack(srcdir, out):
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(Path(srcdir).rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(srcdir).as_posix())

def parse(path):
    return etree.parse(str(path))

def ptext(p):
    return "".join(p.xpath(".//w:t/text()", namespaces=NS)).strip()

def has_visual(p):
    return bool(p.xpath(".//w:drawing|.//w:pict|.//a:blip|.//v:imagedata|.//asvg:svgBlip|.//mc:AlternateContent", namespaces=NS))

def find_caption(root, regex):
    for p in root.xpath("//w:body//w:p", namespaces=NS):
        if regex.search(ptext(p)):
            return p
    return None

def previous_visual(caption, limit=12):
    cur = caption.getprevious()
    seen = 0
    while cur is not None and seen < limit:
        if cur.tag == "{%s}p" % NS["w"] and has_visual(cur):
            return cur
        cur = cur.getprevious()
        seen += 1
    return None

def rel_map(rel_tree):
    out = {}
    for rel in rel_tree.getroot():
        rid = rel.get("Id")
        if rid:
            out[rid] = rel
    return out

def next_rid(rel_tree):
    nums = []
    for rel in rel_tree.getroot():
        m = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if m:
            nums.append(int(m.group(1)))
    n = max(nums or [0]) + 1
    while True:
        rid = f"rId{n}"
        if all(x.get("Id") != rid for x in rel_tree.getroot()):
            return rid
        n += 1

def add_image_rel(rel_tree, target):
    rid = next_rid(rel_tree)
    el = etree.Element(REL + "Relationship")
    el.set("Id", rid)
    el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    el.set("Target", target)
    rel_tree.getroot().append(el)
    return rid

def ensure_png_content_type(base_dir):
    path = base_dir / "[Content_Types].xml"
    tree = parse(path)
    root = tree.getroot()
    exists = any(x.get("Extension", "").lower() == "png" for x in root.findall(CT + "Default"))
    if not exists:
        el = etree.Element(CT + "Default")
        el.set("Extension", "png")
        el.set("ContentType", "image/png")
        root.append(el)
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone="yes")

def copy_rel_target(src_dir, base_dir, src_rel, fig_key, index):
    target = src_rel.get("Target")
    mode = src_rel.get("TargetMode")
    if mode == "External":
        return None, None
    src_part = (src_dir / "word" / target).resolve()
    if not src_part.exists():
        print("WARN target missing", target)
        return None, None
    ext = src_part.suffix.lower()
    media_dir = base_dir / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    out_name = f"repaired_{fig_key.replace('-', '_')}_{index}{ext}"
    out_part = media_dir / out_name
    shutil.copy2(src_part, out_part)
    return out_name, ext

def convert_svg(svg_path, png_path):
    import cairosvg
    cairosvg.svg2png(url=str(svg_path), write_to=str(png_path), output_width=None, output_height=None)

def remove_svg_extensions(p):
    # Remove Office SVG extension blocks so Word/WPS uses the raster fallback only.
    for extlst in p.xpath(".//a:blip/a:extLst", namespaces=NS):
        parent = extlst.getparent()
        parent.remove(extlst)
    # If an AlternateContent exists, prefer its Fallback payload to avoid SVG-only Choice branches.
    for ac in list(p.xpath(".//mc:AlternateContent", namespaces=NS)):
        fb = ac.find("{%s}Fallback" % NS["mc"])
        parent = ac.getparent()
        if fb is not None and parent is not None:
            idx = parent.index(ac)
            children = [copy.deepcopy(x) for x in fb]
            parent.remove(ac)
            for ch in reversed(children):
                parent.insert(idx, ch)

def repair_one(src_dir, base_dir, src_doc, base_doc, src_rels, base_rels, fig_key, regex):
    src_cap = find_caption(src_doc.getroot(), regex)
    base_cap = find_caption(base_doc.getroot(), regex)
    if src_cap is None:
        raise RuntimeError(f"{fig_key}: source caption not found")
    if base_cap is None:
        raise RuntimeError(f"{fig_key}: base caption not found")
    src_vis = previous_visual(src_cap)
    if src_vis is None:
        raise RuntimeError(f"{fig_key}: source visual paragraph not found")

    print(f"{fig_key}: source visual text={ptext(src_vis)!r}")
    src_rel_by_id = rel_map(src_rels)
    copied = copy.deepcopy(src_vis)

    # detect SVG relation if present
    svg_ids = []
    for n in copied.xpath(".//asvg:svgBlip", namespaces=NS):
        rid = n.get(R + "embed")
        if rid:
            svg_ids.append(rid)
    svg_ids = list(dict.fromkeys(svg_ids))
    primary_png_rid = None

    if svg_ids:
        # Render first SVG to PNG and make it the sole raster source.
        svg_rel = src_rel_by_id.get(svg_ids[0])
        if svg_rel is None:
            raise RuntimeError(f"{fig_key}: svg relationship {svg_ids[0]} missing")
        target = svg_rel.get("Target")
        svg_path = (src_dir / "word" / target).resolve()
        media_dir = base_dir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        png_name = f"repaired_{fig_key.replace('-', '_')}.png"
        png_path = media_dir / png_name
        print(f"{fig_key}: converting SVG {target} -> media/{png_name}")
        convert_svg(svg_path, png_path)
        primary_png_rid = add_image_rel(base_rels, "media/" + png_name)

        # Route all standard raster blips / VML image ids to the converted PNG.
        for n in copied.xpath(".//a:blip", namespaces=NS):
            if n.get(R + "embed") is not None:
                n.set(R + "embed", primary_png_rid)
            if n.get(R + "link") is not None:
                del n.attrib[R + "link"]
        for n in copied.xpath(".//v:imagedata", namespaces=NS):
            if n.get(R + "id") is not None:
                n.set(R + "id", primary_png_rid)
        remove_svg_extensions(copied)
    else:
        # Copy every image relationship used by the visual paragraph.
        attrs = []
        for el in copied.iter():
            for name, value in list(el.attrib.items()):
                if name in (R+"embed", R+"link", R+"id") and value:
                    attrs.append((el, name, value))
        mapping = {}
        seq = 1
        for _, _, oldrid in attrs:
            if oldrid in mapping:
                continue
            srel = src_rel_by_id.get(oldrid)
            if srel is None:
                continue
            typ = srel.get("Type", "")
            if not typ.endswith("/image"):
                print(f"{fig_key}: non-image rel skipped {oldrid} {typ}")
                continue
            name, ext = copy_rel_target(src_dir, base_dir, srel, fig_key, seq)
            seq += 1
            if not name:
                continue
            if ext == ".svg":
                svg_path = base_dir / "word" / "media" / name
                png_name = Path(name).with_suffix(".png").name
                png_path = base_dir / "word" / "media" / png_name
                convert_svg(svg_path, png_path)
                name = png_name
            mapping[oldrid] = add_image_rel(base_rels, "media/" + name)
        for el, name, oldrid in attrs:
            if oldrid in mapping:
                el.set(name, mapping[oldrid])
        remove_svg_extensions(copied)

    # Remove blank/old visual immediately before caption in base.
    old = previous_visual(base_cap, limit=5)
    if old is not None:
        print(f"{fig_key}: removing existing visual before caption")
        old.getparent().remove(old)

    # Insert repaired visual immediately before caption.
    base_cap.addprevious(copied)
    print(f"{fig_key}: inserted repaired visual")

def validate(out):
    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        repaired = sorted(x for x in names if x.startswith("word/media/repaired_"))
        print("repaired media:", repaired)
        for n in repaired:
            print(n, len(z.read(n)), "bytes")
        xml = etree.fromstring(z.read("word/document.xml"))
        for key, regex in CAPTIONS:
            cap = find_caption(xml, regex)
            vis = previous_visual(cap) if cap is not None else None
            if vis is None:
                raise RuntimeError(f"validation failed: {key} has no visual")
    print("VALIDATION_OK")

def main():
    if not BASE.exists() or not SOURCE.exists():
        raise SystemExit(f"missing input: BASE={BASE.exists()} SOURCE={SOURCE.exists()}")
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        b = td / "base"; s = td / "src"
        unzip(BASE, b); unzip(SOURCE, s)
        base_doc_path = b / "word/document.xml"
        src_doc_path = s / "word/document.xml"
        base_rels_path = b / "word/_rels/document.xml.rels"
        src_rels_path = s / "word/_rels/document.xml.rels"
        base_doc = parse(base_doc_path)
        src_doc = parse(src_doc_path)
        base_rels = parse(base_rels_path)
        src_rels = parse(src_rels_path)

        print("BASE", BASE, BASE.stat().st_size)
        print("SOURCE", SOURCE, SOURCE.stat().st_size)
        for key, regex in CAPTIONS:
            repair_one(s, b, src_doc, base_doc, src_rels, base_rels, key, regex)

        ensure_png_content_type(b)
        base_doc.write(str(base_doc_path), xml_declaration=True, encoding="UTF-8", standalone="yes")
        base_rels.write(str(base_rels_path), xml_declaration=True, encoding="UTF-8", standalone="yes")
        repack(b, OUT)

    print("OUTPUT", OUT, OUT.stat().st_size)
    validate(OUT)

if __name__ == "__main__":
    main()
