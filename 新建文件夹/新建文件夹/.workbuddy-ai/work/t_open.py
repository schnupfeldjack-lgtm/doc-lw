import win32com.client as win32, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
files = [BASE + r"\炎黄职业技术学院毕业论文模板(1).docx",
 BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx",
 BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx",
 BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx"]
w = win32.DispatchEx("Word.Application"); w.Visible=False; w.DisplayAlerts=0
for f in files:
    try:
        d = w.Documents.Open(f, ReadOnly=True)
        print("OK  ", os.path.basename(f)[:28], "段数=", d.Paragraphs.Count, "页数=", d.ComputeStatistics(2))
        d.Close(False)
    except Exception as e:
        print("FAIL", os.path.basename(f)[:28], e)
w.Quit()
