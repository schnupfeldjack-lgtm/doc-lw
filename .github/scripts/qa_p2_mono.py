# -*- coding: utf-8 -*-
from pathlib import Path
import re,subprocess,sys,zipfile
from docx import Document
from docx.oxml.ns import qn
import pdfplumber

def norm(s):return re.sub(r"\s+","",s).replace("–","-").replace("—","-")
def ptext(pdf,i):return subprocess.check_output(["pdftotext","-f",str(i),"-l",str(i),"-layout",str(pdf),"-"],text=True,errors="ignore")
def pages(pdf):
    s=subprocess.check_output(["pdfinfo",str(pdf)],text=True,errors="ignore")
    return int(re.search(r"^Pages:\s+(\d+)",s,re.M).group(1))

def main():
    docx=Path(sys.argv[1]);pdf=Path(sys.argv[2]);d=Document(docx)
    assert len(d.inline_shapes)>=5
    assert len(d.tables)>=8
    refs=[p.text.strip() for p in d.paragraphs if re.match(r"^\[\d+\]",p.text.strip())]
    assert len(refs)==15,len(refs)
    # 所有正文新增表格禁止底色和有色边框
    for t in d.tables[5:]:
        xml=t._tbl.xml.lower()
        assert 'w:fill="dce6f1"' not in xml and 'w:fill="f7f9fb"' not in xml
        for bad in ("5b9bd5","70ad47","ed7d31","8064a2","2f8f9d","c94c4c"):
            assert bad not in xml
    # 图片二进制真实嵌入
    with zipfile.ZipFile(docx) as z:
        pngs=[x for x in z.namelist() if x.startswith("word/media/") and x.endswith(".png")]
        sizes=[len(z.read(x)) for x in pngs]
        assert sum(1 for x in sizes if x>100000)>=5,(pngs,sizes)
    total=pages(pdf);texts=[ptext(pdf,i) for i in range(1,total+1)];np=[norm(x) for x in texts]
    bphys=next(i for i,t in enumerate(np,1) if "1引言" in t and "市政综合管廊具有线路长" in t)
    assert bphys==9,bphys
    for i in range(1,bphys):
        assert not re.search(r"第\s*\d+\s*页\s*共\s*\d+\s*页",texts[i-1]),i
    body_pages=total-bphys+1
    assert f"第1页共{body_pages}页" in np[bphys-1]
    assert f"第{body_pages}页共{body_pages}页" in np[-1]
    for i,t in enumerate(np,1):assert len(t)>20,f"blank {i}"
    # 每个图题所在页必须检测到一个主要栅格图，且图形必须处于正文页框内部，不能“飞出”。
    with pdfplumber.open(pdf) as f:
        for label in ("图1-1","图2-1","图3-1","图4-1","图5-1"):
            phys=next((i for i,t in enumerate(np,1) if norm(label) in t),None)
            assert phys is not None,label
            pg=f.pages[phys-1]
            ims=pg.images
            assert ims,(label,phys)
            # 最大图片应落在页面主体中，左右留白至少40pt，上下不越界
            im=max(ims,key=lambda x:(x.get("x1",0)-x.get("x0",0))*(x.get("y1",0)-x.get("y0",0)))
            x0,x1=im.get("x0",0),im.get("x1",0)
            top=im.get("top",0);bottom=im.get("bottom",pg.height)
            assert x0>=35 and x1<=pg.width-35,(label,phys,x0,x1,pg.width)
            assert top>=45 and bottom<=pg.height-45,(label,phys,top,bottom,pg.height)
            # 图片不应占据整页造成巨大空白
            assert (bottom-top)<=310,(label,phys,bottom-top)
    print("MONO_QA_OK","pages",total,"body_start",bphys,"body_pages",body_pages,"pics",len(d.inline_shapes),"tables",len(d.tables),"refs",len(refs))

if __name__=="__main__":main()
