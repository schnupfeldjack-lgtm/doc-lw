import win32com.client as win32
f = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹/.workbuddy-ai/work/test_nodeid.docx"
w = win32.DispatchEx("Word.Application"); w.Visible=False; w.DisplayAlerts=0
try:
    d = w.Documents.Open(f, ReadOnly=True)
    print("OK 页数=", d.ComputeStatistics(2), "段数=", d.Paragraphs.Count)
    d.Close(False)
except Exception as e:
    print("FAIL", e)
w.Quit()
