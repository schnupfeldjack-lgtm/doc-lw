# -*- coding: utf-8 -*-
"""修复 fill_docx 临时文档图片关系：把4幅原创图真正嵌入最终 DOCX。"""
from pathlib import Path
import re, shutil, tempfile, zipfile
from lxml import etree

ROOT=Path(__file__).resolve().parents[2]
PAPER=ROOT/"新建文件夹/新建文件夹/再生混凝土在装配式建筑中的应用评价——以住宅项目为例"
DOCX=PAPER/"再生混凝土在装配式建筑中的应用评价——以住宅项目为例_最终优化版.docx"
FIGDIR=ROOT/"新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p1_v2"

W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A="http://schemas.openxmlformats.org/drawingml/2006/main"
REL="http://schemas.openxmlformats.org/package/2006/relationships"
NS={"w":W,"a":A,"r":R}

FIGS=[
 ("图2-1再生骨料特征向预制生产风险的传导路径","fig2_1.png"),
 ("图3-1百米级再生混凝土工程的可核验事实链","fig3_1.png"),
 ("图4-1再生混凝土向装配式住宅迁移的证据分层矩阵","fig4_1.png"),
 ("图5-1装配式再生混凝土的质量闸门与放行流程","fig5_1.png"),
]

def norm(s):return re.sub(r"\s+","",s).replace("–","-").replace("—","-")
def text(el):return "".join(el.xpath(".//w:t/text()",namespaces=NS)).strip()

def next_rids(relroot,n):
    nums=[]
    for r in relroot:
        m=re.fullmatch(r"rId(\d+)",r.get("Id",""))
        if m:nums.append(int(m.group(1)))
    start=max(nums or [0])+1
    return [f"rId{start+i}" for i in range(n)]

def main():
    if not DOCX.exists():raise FileNotFoundError(DOCX)
    for _,fn in FIGS:
        if not (FIGDIR/fn).exists():raise FileNotFoundError(FIGDIR/fn)
    with tempfile.TemporaryDirectory() as td:
        td=Path(td)
        with zipfile.ZipFile(DOCX) as z:z.extractall(td)
        dp=td/"word/document.xml";rp=td/"word/_rels/document.xml.rels";media=td/"word/media"
        media.mkdir(parents=True,exist_ok=True)
        dt=etree.parse(str(dp));rt=etree.parse(str(rp));root=dt.getroot();rels=rt.getroot()
        rids=next_rids(rels,len(FIGS))
        for (prefix,fn),rid in zip(FIGS,rids):
            cap=None
            for p in root.xpath("//w:body//w:p",namespaces=NS):
                if norm(text(p)).startswith(prefix):
                    cap=p;break
            if cap is None:raise RuntimeError("找不到图题 "+prefix)
            cur=cap.getprevious();pic=None
            for _ in range(8):
                if cur is None:break
                if cur.tag==f"{{{W}}}p" and cur.xpath(".//w:drawing",namespaces=NS):
                    pic=cur;break
                cur=cur.getprevious()
            if pic is None:raise RuntimeError("图题前没有图片段 "+prefix)
            blips=pic.xpath(".//a:blip",namespaces=NS)
            if not blips:raise RuntimeError("图片没有 a:blip "+prefix)
            outname="p1v2_"+fn
            shutil.copy2(FIGDIR/fn,media/outname)
            rel=etree.Element(f"{{{REL}}}Relationship")
            rel.set("Id",rid)
            rel.set("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            rel.set("Target","media/"+outname)
            rels.append(rel)
            for b in blips:
                b.set(f"{{{R}}}embed",rid)
                if f"{{{R}}}link" in b.attrib:del b.attrib[f"{{{R}}}link"]
            print("PATCHED",prefix,rid,outname,(media/outname).stat().st_size)
        dt.write(str(dp),encoding="UTF-8",xml_declaration=True,standalone="yes")
        rt.write(str(rp),encoding="UTF-8",xml_declaration=True,standalone="yes")
        tmp=DOCX.with_suffix(".patched.docx")
        with zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as z:
            for p in td.rglob("*"):
                if p.is_file():z.write(p,p.relative_to(td).as_posix())
        shutil.move(tmp,DOCX)
    with zipfile.ZipFile(DOCX) as z:
        for _,fn in FIGS:
            name="word/media/p1v2_"+fn
            assert name in z.namelist() and len(z.read(name))>50000,name
    print("IMAGE_PATCH_OK",DOCX,DOCX.stat().st_size)

if __name__=="__main__":main()
