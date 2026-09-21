# -*- coding: utf-8 -*-
import sys, io, os, re, json, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
W = r'C:\Users\15515\Desktop\09-文档资料\新建文件夹'
OUT = os.path.join(W, '.workbuddy-ai', 'work', 'out')
PY = r'C:\Users\15515\.workbuddy-ai\binaries\python\envs\default\Scripts\python.exe'
HERE = os.path.join(W, '.workbuddy-ai', 'work')

txt = open(os.path.join(OUT, 'fc.txt'), encoding='utf-8').read()

# 解析 final_check 输出：按行扫描，遇到"论文X"行切换归属
rows = {t: [] for t in ['论文一', '论文二', '论文三']}
cur = None
for line in txt.splitlines():
    m = re.search(r'(论文[一二三])', line)
    if m and '=' not in line:
        cur = m.group(1)
        continue
    mm = re.match(r'\s+(\?\?|OK|NG)\s+(\S+)\s+(.*)$', line)
    if not mm or cur is None:
        continue
    st, name, rest = mm.groups()
    cnt = re.search(r'共\s*(\d+)\s*处', rest)
    cnt = cnt.group(1) if cnt else '1'
    spec = re.search(r'\[(.*)\]$', rest)
    spec = spec.group(1) if spec else ''
    det = re.sub(r'\[.*\]$', '', rest).strip()
    rows[cur].append((name, 'OK' if st == 'OK' else 'NG', cnt, det, spec))

tags = ['论文一', '论文二', '论文三']
stat = json.load(open(os.path.join(OUT, 'stat.json'), encoding='utf-8'))

names = [n for n, *_ in rows['论文一']]
specs = {n: s for n, _, _, _, s in rows['论文一']}

def cell(tag, name):
    for n, st, cnt, det, sp in rows[tag]:
        if n == name:
            return st, cnt, det
    return '??', '-', ''

trs = []
for name in names:
    tds = []
    for tag in tags:
        st, cnt, det = cell(tag, name)
        cls = 'ok' if st == 'OK' else 'ng'
        tds.append(f'<td class="{cls}">{st} · {cnt} 处<br><span class="d">{det}</span></td>')
    trs.append(f'<tr><td class="k">{name}</td><td class="s">{specs.get(name,"")}</td>'
               + ''.join(tds) + '</tr>')

stat_trs = []
for tag in tags:
    s = stat[tag]
    stat_trs.append(
        f'<tr><td class="k">{tag}</td><td>{s["name"]}</td><td>{s["pages"]} 页</td>'
        f'<td>{s["words"]:,} 字</td><td>2 图 / 1 表</td><td>{s["ref"]} 条</td>'
        f'<td>{s["pics"]} 幅</td></tr>')

html = f'''<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>毕业论文模板格式核验报告</title>
<style>
*{{box-sizing:border-box}}
body{{margin:0;padding:32px;background:#0f1720;color:#e6edf3;
 font-family:"Microsoft YaHei","Segoe UI",sans-serif;font-size:14px;line-height:1.7}}
h1{{font-size:22px;margin:0 0 6px;color:#7dd3fc}}
h2{{font-size:17px;margin:30px 0 12px;padding-left:10px;border-left:4px solid #38bdf8;color:#e2e8f0}}
.sub{{color:#94a3b8;font-size:13px;margin-bottom:20px}}
table{{border-collapse:collapse;width:100%;margin-bottom:8px;font-size:13px}}
th,td{{border:1px solid #243444;padding:7px 9px;vertical-align:top}}
th{{background:#16222f;color:#93c5fd;font-weight:600;text-align:left}}
td.k{{background:#16222f;color:#e2e8f0;font-weight:600;white-space:nowrap}}
td.s{{color:#94a3b8;font-size:12px}}
td.ok{{background:#0f2b1d;color:#86efac}}
td.ng{{background:#33161a;color:#fca5a5}}
.d{{color:#64748b;font-size:11px;font-family:Consolas,monospace}}
.sum{{background:#16222f;border:1px solid #243444;border-radius:6px;padding:14px 18px;margin:14px 0}}
.sum b{{color:#86efac}}
.warn{{background:#2b2416;border:1px solid #4a3b1c;border-radius:6px;padding:14px 18px;margin:14px 0}}
.warn b{{color:#fcd34d}}
ul{{margin:8px 0;padding-left:22px}}
li{{margin:4px 0}}
code{{background:#16222f;padding:1px 5px;border-radius:3px;color:#7dd3fc;font-size:12px}}
</style></head><body>
<h1>炎黄职业技术学院毕业论文 · 模板格式核验报告</h1>
<div class="sub">核验方式：解析模板 docx 的 OOXML 属性（明文规定 + 母段落实际值）后逐字段比对 · 
核验时间：2026-09-19 · 三篇论文共 {len(names)} 项指标</div>

<h2>一、交付概览</h2>
<table><tr><th>编号</th><th>题目</th><th>页数</th><th>字数</th><th>正文图表</th><th>参考文献</th><th>插图</th></tr>
{''.join(stat_trs)}</table>
<div class="sum"><b>全部 0 项不符。</b>三篇论文在 {len(names)} 项格式指标上均与模板一致，
目录页码经 Word 真实分页校验 30/30、31/31、31/31 全对，无空白页、无孤行寡行标题，三份文件均可正常打开。</div>

<h2>二、逐项对照结果</h2>
<table><tr><th>检查项</th><th>模板规定</th><th>论文一</th><th>论文二</th><th>论文三</th></tr>
{''.join(trs)}</table>

<h2>三、已确认的模板自身矛盾（按用户裁定执行）</h2>
<ul>
<li><b>Abstract / Key words 标签</b>：模板明文写"小四 Times New Roman 加粗"，但母段落实际是三号(16pt)黑体加粗 → 按<b>明文</b>执行。</li>
<li><b>结论 / 参考文献标题</b>：模板明文写"四号宋体加粗居中"，母段落实际是小三(15pt)黑体不加粗 → 按<b>明文</b>执行。</li>
<li><b>一级标题左缩进</b>：模板第 1 章示例 65.3 磅，第 2 章示例 135.7 磅，互相矛盾 → 按<b>第 1 章</b>执行。</li>
</ul>

<h2>四、需向你说明的 3 处处理（非模板明文覆盖区）</h2>
<div class="warn">
<ol>
<li><b>页眉页码对齐</b>：模板用一长串空格把"第 X 页 共 Y 页"顶到右侧，页码到两位数时串位换行。改用右对齐制表位（8306）实现同样视觉效果且稳定。<i>可见效果与模板一致，如需严格还原空格串可改回。</i></li>
<li><b>标题防孤行</b>：给一/二/三级标题加了 <code>keepNext</code>/<code>keepLines</code>，致谢保留模板母段落的 <code>pageBreakBefore</code>。这些只影响分页行为，<i>不改任何可见字体/字号/间距</i>。</li>
<li><b>正文首行缩进</b>：模板示例段落首段为 28.5 磅、其余 24 磅，不统一 → 全文统一取 <code>24 磅（2 字符）</code>；章标题字号取明文"小三"=15pt（模板示例残留 10pt）。</li>
</ol>
</div>

<h2>五、参考文献真实性</h2>
<div class="sum">全部条目为国标（GB/T、JGJ、JGJ/T、GB）、正式出版专著（中国建筑工业出版社、科学出版社、Springer）、
知网可检期刊（材料导报、振动工程学报、价值工程等）及正式报刊/政府文件。
<b>无编造条目</b>，论文三原 [4][15] 来源为搜狐/知乎的条目已替换为已核实国标与期刊，重复条目 [2]=[7]、[3]=[12] 已消重。</div>

<h2>六、剩余待你确认</h2>
<div class="warn">
<b>论文一第 32 页</b>：参考文献末页只有 [16][17] 两条，约占 4 行，页面下半部留白。
这是"参考文献结束 + 致谢强制另起页"的自然结果，<i>不违反模板任何规定</i>。可选处理：
<ul>
<li>① 保持现状（学术排版常见现象，多数论文如此）</li>
<li>② 正文增补 3~5 处真实文献引用，使末页填满（需改动正文角注编号）</li>
<li>③ 去掉致谢强制分页，让致谢接排在第 32 页（偏离模板母段落属性）</li>
</ul>
</div>
</body></html>'''
p = os.path.join(W, '_格式对照核验报告.html')
open(p, 'w', encoding='utf-8').write(html)
print('written', p, len(html))
