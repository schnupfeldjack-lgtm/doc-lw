# -*- coding: utf-8 -*-
"""结论页处理：删除结论前独立的分页符段落内容（保留为空行以凑"空2行"），
   改为给"结  论"标题设段前分页（w:pageBreakBefore），符合模板"本页为独立页"且不产生空白页"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
FILES = ['再生混凝土在装配式建筑中的应用评价——以住宅项目为例',
         '装配式施工质量管理问题及优化研究——以市政项目为例',
         'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例']

for sub in FILES:
    p_ = os.path.join(BASE, sub, sub+".docx")
    d = docx.Document(p_); ps = d.paragraphs
    for i, p in enumerate(ps):
        if p.text.strip() == '结  论':
            # 1) 前一段若是分页符段落 -> 清空为空段落
            prev = ps[i-1]
            if '\x0c' in prev.text or '\f' in prev.text:
                for r in list(prev.runs):
                    r._element.getparent().remove(r._element)
                print(f"  清空分页符段[{i-1}]")
            # 2) 结论标题设段前分页
            pPr = p._element.get_or_add_pPr()
            if pPr.find(qn('w:pageBreakBefore')) is None:
                pb = OxmlElement('w:pageBreakBefore')
                pPr.append(pb)
                print(f"  结论段[{i}] 已设 pageBreakBefore")
            # 确认前面有 2 个空段
            blanks = 0; j = i-1
            while j >= 0 and not ps[j].text.strip():
                blanks += 1; j -= 1
            print(f"  结论前空段数={blanks}")
    d.save(p_)
    print(f"{sub[:24]} 已保存")
