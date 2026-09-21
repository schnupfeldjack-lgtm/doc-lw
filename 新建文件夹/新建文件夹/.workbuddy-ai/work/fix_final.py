# -*- coding: utf-8 -*-
"""
按"模板明文"修正最后一批偏差：
1) 摘要：标签「摘 要：」加粗、正文不加粗；正文行距改 1.5 倍（明文"小四号宋体，1.5倍行距"）；字数压到 ≤300
2) 章标题：段前 15.6pt(1行) → 7.8pt（明文"段前0.5行"，与模板第2章母段落一致）
3) 正文：首行缩进 28.5pt → 24pt（模板正文母段落 480twips=2字符）
4) 图/表题：首行缩进清零，保证真正居中
5) 正文行文中的"图3-1/表4-1"半角连字符 → en dash，与图题编号一致（明文"图3–5"）
"""
import re, copy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f"{{{W}}}{t}"

BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
DOCS = [
    (BASE + r"\再生混凝土在装配式建筑中的应用评价——以住宅项目为例\再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx", "论文一",
     "再生混凝土由废弃混凝土破碎、筛分后作骨料重新拌制而成，可缓解建筑固废堆存与天然砂石资源紧张。本文以广西某30层、高度接近100 m的再生混凝土公寓项目为对象，对其再生骨料取代率30%、总用量约10679 m³的应用情况进行评价。研究从技术性能、施工适应性、经济与环境效益三个维度构建评价指标体系，结合配合比设计、构件生产与实体检测数据开展分析。结果表明：30%取代率下再生混凝土的抗压强度、弹性模量与耐久性能均满足设计要求，弹性模量略低有利于优化结构刚度分布；但骨料吸水率偏高引起的工作性波动与构件质量离散性偏大仍需控制。对此提出加强骨料分级预处理、优化配合比参数、完善过程质量留痕等对策。"),
    (BASE + r"\装配式施工质量管理问题及优化研究——以市政项目为例\装配式施工质量管理问题及优化研究——以市政项目为例.docx", "论文二",
     "装配式建造可把大量湿作业转移到工厂，减少现场用工与城市干扰，是综合管廊工业化建造的重要方向。本文以绵阳综合管廊项目为对象，针对管廊总长约25.15 km、配套市政道路27.14 km的工程实际，围绕装配式施工质量问题开展调查与成因分析并给出优化对策。研究从人、机、料、法、环五维梳理构件进场验收、吊装定位、接缝防水、套筒连接和多专业穿插等典型质量问题及其成因，进而从组织管理、工艺技术和质量控制手段三方面提出措施，包括建立构件批次追溯制度、推行驻厂监理、应用BIM与三维扫描相结合的偏差动态修正工艺、采用预铺反粘与自愈合高分子防水卷材复合工艺等。相关对策已在项目应用，预埋件安装合格率达到100%。"),
    (BASE + r"\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例\BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx", "论文三",
     "BIM协同设计是提升装配式建筑结构设计质量的重要技术手段。本文以上海闵行浦江镇召楼路以东S8-01市属保障房项目为对象，围绕总建筑面积约11万㎡、采用装配整体式剪力墙体系且装配式比例100%的工程实际，分析BIM协同设计的应用效果。研究从设计协同、构件深化、管线综合和设计变更四个维度分析其影响机理，结合参数化构件库构建、碰撞检查、节点优化和设计—生产数据贯通等实践，探讨其在减少构件尺寸错误、预埋件位置偏差、管线与结构冲突等问题的效果。结果表明，BIM协同设计可显著提高设计一次性成品率。针对协同深度不足、数据传递不连续等问题，提出加强全专业协同、推进设计—生产一体化、加强人员培训等对策。"),
]

def get_or(el, path):
    cur = el
    for tag in path:
        nxt = cur.find(q(tag))
        if nxt is None:
            nxt = OxmlElement('w:' + tag)
            cur.append(nxt)
        cur = nxt
    return cur

def set_spacing(pPr, **kw):
    sp = get_or(pPr, ['spacing'])
    for k, v in kw.items():
        sp.set(q(k), v)

for path, nm, new_abs in DOCS:
    d = Document(path)
    n_abs = n_h1 = n_body = n_cap = n_dash = 0

    for p in d.paragraphs:
        t = p.text.strip()
        pPr = p._p.find(q('pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr'); p._p.insert(0, pPr)
        ind = pPr.find(q('ind'))
        fl = ind.get(q('firstLine')) if ind is not None else None
        jc = pPr.find(q('jc'))
        jcv = jc.get(q('val')) if jc is not None else None

        # --- 1 摘要 ---
        if re.match(r'^摘\s*要[：:]', t):
            run = p.runs[0]
            rr = run._element.find(q('rPr'))
            run._element.find(q('t')).text = '摘 要：'
            # 正文 run：去掉加粗
            r2 = OxmlElement('w:r')
            if rr is not None:
                nr = copy.deepcopy(rr)
                for b in nr.findall(q('b')) + nr.findall(q('bCs')):
                    nr.remove(b)
                r2.append(nr)
            t2 = OxmlElement('w:t')
            t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t2.text = new_abs
            r2.append(t2)
            run._element.addnext(r2)
            set_spacing(pPr, line='360', lineRule='auto')
            n_abs += 1
            continue

        # --- 2 章标题 段前 0.5 行 ---
        if re.match(r'^\d+\s{1,2}\S', t) and '\t' not in t and len(t) < 60 and not re.match(r'^\d+\.\d', t) and not re.match(r'^\d+．\d', t):
            set_spacing(pPr, before='156', beforeLines='50', after='156', afterLines='50')
            n_h1 += 1
            continue

        # --- 4 图/表题 缩进清零 ---
        if re.match(r'^(图|表)\s*\d+[–\-]\d+\s{2}\S', t) and len(t) < 50:
            if ind is not None and ind.get(q('firstLine')):
                ind.set(q('firstLine'), '0')
                if ind.get(q('firstLineChars')):
                    ind.set(q('firstLineChars'), '0')
                n_cap += 1
            continue

        # --- 3 正文首行缩进 570 -> 480 ---
        if fl == '570' or fl == '571':
            ind.set(q('firstLine'), '480')
            if ind.get(q('firstLineChars')):
                ind.set(q('firstLineChars'), '200')
            n_body += 1

        # --- 5 行文中的 图3-1 / 表4-1 改 en dash ---
        if len(t) > 5:
            for r in p.runs:
                te = r._element.find(q('t'))
                if te is None or not te.text:
                    continue
                nt = re.sub(r'(图|表)(\d+)-(\d+)', r'\1\2–\3', te.text)
                if nt != te.text:
                    te.text = nt
                    n_dash += 1

    d.save(path)
    print(f"{nm}: 摘要修正={n_abs} 章标题段前={n_h1} 正文缩进={n_body} 图题缩进={n_cap} 编号en dash={n_dash}")
print("完成")
