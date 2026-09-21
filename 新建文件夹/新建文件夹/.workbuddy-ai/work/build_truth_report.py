# -*- coding: utf-8 -*-
import io, os
BASE = r"C:\Users\15515\Desktop\09-文档资料\新建文件夹"
HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>独立复核报告（含仍待裁定项）</title>
<style>
body{font-family:"Microsoft YaHei",sans-serif;background:#f6f7fb;color:#1f2430;margin:0;padding:28px}
.wrap{max-width:1060px;margin:0 auto}
h1{font-size:23px;margin:0 0 6px}.sub{color:#6b7280;font-size:13px;margin-bottom:20px}
.card{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:18px 22px;margin-bottom:16px}
h2{font-size:17px;margin:0 0 12px;padding-left:10px;border-left:4px solid #2563eb}
h2.warn{border-left-color:#d97706} h2.todo{border-left-color:#dc2626}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th,td{border:1px solid #e5e7eb;padding:7px 10px;text-align:left;vertical-align:top}
th{background:#f1f5f9;font-weight:600}
.ok{color:#15803d;font-weight:600}.bad{color:#b91c1c;font-weight:600}
.mono{font-family:Consolas,monospace;font-size:12.5px}
code{background:#f1f5f9;padding:1px 5px;border-radius:3px;font-family:Consolas,monospace}
.note{background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:12px 16px;margin-bottom:16px}
</style></head><body><div class="wrap">
<h1>独立复核报告</h1>
<div class="sub">说明：此前报的"0 不符"，只是我自己写的 final_check.py 里 24 项的自证，不能作为"已符合模板"的证据。本轮换成三条独立路径复核。</div>

<div class="note"><b>复核方法</b>：①段落格式指纹比对（论文每个段落的属性组合必须在模板中出现）；②Word COM 读真实分节与分页；③把模板和论文都导出 PDF，实测每页文本的左右边界（cm）逐页对照。</div>

<div class="card"><h2 class="warn">一、查出并已修复的真实问题</h2>
<table><tr><th style="width:24%">问题</th><th>证据</th><th style="width:22%">处理</th></tr>
<tr><td><b>分节数不对，前 4 页页边距错误</b><br>（最严重）</td>
<td>模板 <b>4 节</b>：节1（开题报告/任务书/中期检查表/封面）= 左3.00 右3.00 装订0.20；节2（摘要/Abstract）= 左2.50 右1.80 装订0.90；节3（目录）、节4（正文）= 左3.17 右3.17。<br>
论文原为 <b>3 节</b>，把封面段并进了 2.50/1.80/0.90 那一节。<br>
PDF 实测左边界：模板 p1–p3 = <span class="mono">3.20cm</span>，论文原为 <span class="mono">3.41cm</span>（差 0.21cm）。</td>
<td class="ok">在封面末段插入模板节1 的 sectPr。<br>修复后逐页实测：<br>论文 <span class="mono">3.20 / 3.20 / 3.20 / 4.06 / 3.41 / 3.41 / 3.18…</span><br>模板 <span class="mono">3.20 / 3.20 / 3.20 / 4.06 / 3.41 / 3.41 / 3.18…</span><br>完全一致</td></tr>
<tr><td>页眉页码域缓存是模板残留</td>
<td>页眉缓存值仍为模板的"第 1 页  共 4 页"（模板只有 11 页，论文 30+ 页）</td>
<td class="ok">改为真实页数 32 / 30 / 31</td></tr>
<tr><td class="bad">Word 保存会破坏格式（过程中的坑）</td>
<td>用 Word COM <code>doc.Save()</code> 后，run 的 <code>w:eastAsia</code>（中文字体）和 <code>w:sz</code>（字号）被丢弃，run 被拆碎，final_check 立刻从 0 不符变 7 处不符</td>
<td class="ok">已从备份回退；改为 Word 只读取分页、写入一律用 python-docx</td></tr>
</table></div>

<div class="card"><h2 class="todo">二、仍待你裁定（模板自身不统一或没有明文，我没擅自改）</h2>
<table><tr><th style="width:22%">项目</th><th>模板实际情况</th><th>论文当前</th></tr>
<tr><td>正文首行缩进</td><td>正文示例段落有三种值：<code>570</code>(28.5磅，紧跟一级标题那段)、<code>480</code>(24磅)、<code>573</code>；第 2 章后又变回 480 —— <b>模板自身不统一</b></td><td>全文统一 480（24磅）</td></tr>
<tr><td>页眉"第 X 页 共 Y 页"</td><td>用一长串空格把页码顶到右侧；两位数页码时<b>模板自己就串位换行</b>（模板 p10/p11 可看到"共 11 \n页"）</td><td>改用右对齐制表位，不串位</td></tr>
<tr><td>标题防孤行</td><td>母段落无 keepNext / keepLines</td><td>二、三级标题加了（只在分页时起作用，不改可见格式）</td></tr>
<tr><td>正文表格字号</td><td>模板 5 张表都是封面/任务书类，字号为小四 12pt；<b>没有正文表格示例</b></td><td>正文表格用五号 10.5pt（与"图与表一律 5 号宋体"一致）</td></tr>
<tr><td>目录形式</td><td>是 TOC 域（自动生成）</td><td>静态文本。好处：任何软件打开页码都不会因未更新域而错</td></tr>
<tr><td>图题/表题居中、图题置于图下方</td><td>模板无图题示例，明文只规定"5 号宋体"</td><td>居中，5 号宋体，图题在图下方</td></tr>
<tr><td>结论是否独立成页</td><td>TPL[88] 写"（本页为独立页，空 2 行）"</td><td>已按独立页处理（若学校不要求，可去掉）</td></tr>
</table></div>

<div class="card"><h2>三、修复后的核验结果</h2>
<table><tr><th style="width:32%">核验项</th><th>论文一</th><th>论文二</th><th>论文三</th></tr>
<tr><td>分节数 / 各节页边距</td><td class="ok">4 节 与模板一致</td><td class="ok">4 节 与模板一致</td><td class="ok">4 节 与模板一致</td></tr>
<tr><td>PDF 逐页左边界 vs 模板</td><td class="ok">前 6 页完全一致</td><td class="ok">前 6 页完全一致</td><td class="ok">前 6 页完全一致</td></tr>
<tr><td>24 项格式核验不符数</td><td class="ok">0</td><td class="ok">0</td><td class="ok">0</td></tr>
<tr><td>文献编号连续 / 角注对应</td><td class="ok">17 条 OK</td><td class="ok">12 条 OK</td><td class="ok">15 条 OK</td></tr>
<tr><td>目录页码（Word 真实分页）</td><td class="ok">30/30</td><td class="ok">31/31</td><td class="ok">31/31</td></tr>
<tr><td>空白页</td><td class="ok">0</td><td class="ok">0</td><td class="ok">0</td></tr>
<tr><td>页眉域缓存</td><td class="ok">共 32 页</td><td class="ok">共 30 页</td><td class="ok">共 31 页</td></tr>
<tr><td>总页数</td><td>32 页</td><td>30 页</td><td>31 页</td></tr>
</table></div>
</div></body></html>"""
out = os.path.join(BASE, "_独立复核报告.html")
with io.open(out, "w", encoding="utf-8") as f: f.write(HTML)
print("已生成:", out)
