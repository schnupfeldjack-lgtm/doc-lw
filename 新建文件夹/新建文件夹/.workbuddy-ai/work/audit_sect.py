# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import win32com.client as wc
W = r'C:/Users/15515/Desktop/09-文档资料/新建文件夹'
app = wc.gencache.EnsureDispatch('Word.Application'); app.Visible=False; app.DisplayAlerts=0

def rep(path, tag):
    doc = app.Documents.Open(path, ReadOnly=True); doc.Repaginate()
    print(f"\n===== {tag}  共{doc.ComputeStatistics(2)}页 =====")
    for si in range(1, doc.Sections.Count+1):
        s = doc.Sections(si)
        pf = s.PageSetup
        # 该节第一段的页码
        p0 = s.Range.Paragraphs(1)
        pg = p0.Range.Information(3)
        print(f"  节{si}: 左{pf.LeftMargin/28.35:.2f}cm 右{pf.RightMargin/28.35:.2f}cm "
              f"装订{pf.Gutter/28.35:.2f}cm  起始页≈{pg}  起始段文本={p0.Range.Text[:34]!r}")
    doc.Close(False)

rep(os.path.join(W,'炎黄职业技术学院毕业论文模板(1).docx'), "模板")
for nm in ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
           '装配式施工质量管理问题及优化研究——以市政项目为例',
           'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']:
    rep(os.path.join(W,nm,nm+'.docx'), nm[:20])
app.Quit()
