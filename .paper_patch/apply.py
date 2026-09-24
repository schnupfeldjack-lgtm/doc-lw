# -*- coding: utf-8 -*-
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import base64, bz2, json, re, shutil, tempfile

ROOT = Path("新建文件夹/新建文件夹")
PATCH = Path(".paper_patch")

DOCS = {
    "BIM": {
        "folder": "BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例",
        "file": "BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx",
        "figs": [
            ROOT / ".workbuddy-ai/work/figs_p3/fig3_1.png",
            ROOT / ".workbuddy-ai/work/figs_p3/fig4_1.png",
        ],
        "expect_cn": 8, "expect_en": 7,
    },
    "RAC": {
        "folder": "再生混凝土在装配式建筑中的应用评价——以住宅项目为例",
        "file": "再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
        "figs": [
            ROOT / ".workbuddy-ai/work/figs_p1/fig3_1.png",
            ROOT / ".workbuddy-ai/work/figs_p1/fig4_1.png",
        ],
        "expect_cn": 9, "expect_en": 8,
    },
    "MUNI": {
        "folder": "装配式施工质量管理问题及优化研究——以市政项目为例",
        "file": "装配式施工质量管理问题及优化研究——以市政项目为例.docx",
        "figs": [
            ROOT / ".workbuddy-ai/work/figs_p2/fig3_1.png",
            ROOT / ".workbuddy-ai/work/figs_p2/fig5_1.png",
        ],
        "expect_cn": 6, "expect_en": 6,
    },
}

def read_payload(key):
    chunks = sorted(PATCH.glob(f"{key}_*.txt"))
    if not chunks:
        raise RuntimeError(f"missing patch chunks for {key}")
    return json.loads("".join(p.read_text(encoding="utf-8") for p in chunks))

def replace_docx(src: Path, payload, figs):
    changes = {n: bz2.decompress(base64.b64decode(v)) for n, v in payload["changed"].items()}
    changes["word/media/paper_figure_1.png"] = figs[0].read_bytes()
    changes["word/media/paper_figure_2.png"] = figs[1].read_bytes()
    removed = set(payload.get("removed", []))
    tmp = src.with_suffix(".tmp.docx")
    seen = set()
    with ZipFile(src, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            name = info.filename
            if name in removed:
                continue
            data = changes.get(name, zin.read(name))
            zout.writestr(info, data)
            seen.add(name)
        for name, data in changes.items():
            if name not in seen:
                zout.writestr(name, data)
    with ZipFile(tmp) as z:
        bad = z.testzip()
        if bad:
            raise RuntimeError(f"corrupt zip member {bad}")
    tmp.replace(src)

def document_text(docx: Path):
    with ZipFile(docx) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    texts = re.findall(r"<w:t(?: [^>]*)?>(.*?)</w:t>", xml)
    import html
    return "".join(html.unescape(t) for t in texts)

for key, cfg in DOCS.items():
    doc = ROOT / cfg["folder"] / cfg["file"]
    if not doc.exists():
        raise FileNotFoundError(doc)
    for fig in cfg["figs"]:
        if not fig.exists():
            raise FileNotFoundError(fig)
    payload = read_payload(key)
    replace_docx(doc, payload, cfg["figs"])
    txt = document_text(doc)
    if "参考文献" not in txt or "致谢" not in txt:
        raise RuntimeError(f"{key}: required thesis sections missing")
    if "[J]" not in txt:
        raise RuntimeError(f"{key}: journal references missing")
    if any(x in txt for x in ["住房和城乡建设部.", "国务院."]):
        raise RuntimeError(f"{key}: non-journal policy reference residue found")
    print(f"updated {key}: {doc}")

print("three thesis DOCX files updated successfully")
