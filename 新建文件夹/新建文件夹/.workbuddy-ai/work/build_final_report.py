# -*- coding: utf-8 -*-
import io, os
BASE = r"C:/Users/15515/Desktop/09-文档资料/新建文件夹"
HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>毕业论文格式终版核验报告</title>
<style>
body{font-family:"Microsoft YaHei",sans-serif;background:#f6f7fb;color:#1f2430;margin:0;padding:28px}
.wrap{max-width:1080px;margin:0 auto}
h1{font-size:24px;margin:0 0 6px}
.sub{color:#6b7280;font-size:13px;margin-bottom:22px}
.card{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:18px 22px;margin-bottom:16px}
h2{font-size:17px;margin:0 0 12px;padding-left:10px;border-left:4px solid #2563eb}
table{width:100%;border-collapse:collapse;font-size:13.5px}
th,td{border:1px solid #e5e7eb;padding:7px 10px;text-align:left;vertical-align:top}
th{background:#f1f5f9;font-weight:600}
.ok{color:#15803d;font-weight:600}
.ng{color:#b91c1c;font-weight:600}
.tag{display:inline-block;background:#dcfce7;color:#166534;border-radius:4px;padding:1px 7px;font-size:12px}
code{background:#f1f5f9;padding:1px 5px;border-radius:3px;font-family:Consolas,monospace}
.mono{font-family:Consolas,monospace;font-size:13px}
</style></head><body><div class="wrap">
<h1>炎黄职业技术学院毕业论文 · 格式终版核验报告</h1>
<div class="sub">核验对象：3 篇论文 ｜ 依据：模板明文规定 + 模板母段落实测值 ｜ 工具：python-docx / Word COM 真实分页</div>

<div class="card"><h2>一、本轮修复的 6 项问题</h2>
<table><tr><th style="width:38%">问题</th><th>模板依据 / 原因</th><th style="width:24%">处理</th></tr>
<tr><td>多文献角注写成连写 <code>[15][16]</code></td>
<td>模板明文："应用两篇及以上文献论述同一观点时……如：××××××[4，5]；×××××[6~8]"</td>
<td class="ok">合并为 <code>[15，16]</code></td></tr>
<tr><td>论文一 段179 <code>[5][15][16]</code></td>
<td>三篇共同论述"标准体系现状"这一观点</td>
<td class="ok">合并为 <code>[5，15，16]</code></td></tr>
<tr><td>论文三 段104 语句被角注截断</td>
<td>"…作了系统规定[3][12]预制装配式…"缺标点、[12]位置错</td>
<td class="ok">断句并归位：<code>[3]。…提供了参考[12]。</code></td></tr>
<tr><td>论文一 图4–1 正文无引用句</td>
<td>模板注1："图与表应设置在文章中首次提到处附近"</td>
<td class="ok">补"…如图4–1所示。"</td></tr>
<tr><td>结论前残留独立分页符 → 整页空白</td>
<td>模板 TPL[88]"（本页为独立页，空2行）"</td>
<td class="ok">分页改设于"空2行"首段，结论独立页且无空白页</td></tr>
<tr><td>目录页码错乱 / 末页寡行</td>
<td>目录须与实际分页一致</td>
<td class="ok">Word 真实分页重算；致谢设段中不分页</td></tr>
</table></div>

<div class="card"><h2>二、核验结果（全部通过）</h2>
<table><tr><th style="width:34%">核验项</th><th>论文一</th><th>论文二</th><th>论文三</th></tr>
<tr><td>24 项格式核验不符数</td><td class="ok">0</td><td class="ok">0</td><td class="ok">0</td></tr>
<tr><td>文献编号连续 1..N</td><td class="ok">17 条 OK</td><td class="ok">12 条 OK</td><td class="ok">15 条 OK</td></tr>
<tr><td>角注↔文献 一一对应</td><td class="ok">无未引/无悬空</td><td class="ok">无未引/无悬空</td><td class="ok">无未引/无悬空</td></tr>
<tr><td>角注首次出现顺序递增</td><td class="ok">OK</td><td class="ok">OK</td><td class="ok">OK</td></tr>
<tr><td>题名/摘要/关键词/目录角注</td><td class="ok">0 处</td><td class="ok">0 处</td><td class="ok">0 处</td></tr>
<tr><td>摘要区 5 处空行规定</td><td class="ok">全 OK</td><td class="ok">全 OK</td><td class="ok">全 OK</td></tr>
<tr><td>结论/参考文献/致谢前空2行</td><td class="ok">OK</td><td class="ok">OK</td><td class="ok">OK</td></tr>
<tr><td>目录页码（Word 真实分页）</td><td class="ok">30/30</td><td class="ok">31/31</td><td class="ok">31/31</td></tr>
<tr><td>空白页</td><td class="ok">0</td><td class="ok">0</td><td class="ok">0</td></tr>
<tr><td>模板残留 / 彩色字 / "打印删除"</td><td class="ok">0</td><td class="ok">0</td><td class="ok">0</td></tr>
<tr><td>图表在首次提及处附近</td><td class="ok">3/3</td><td class="ok">3/3</td><td class="ok">3/3</td></tr>
<tr><td>总页数 / 字数</td><td>31 页</td><td>29 页</td><td>30 页</td></tr>
</table></div>

<div class="card"><h2>三、关键方法论（避免再踩坑）</h2>
<table>
<tr><th style="width:30%">坑</th><th>说明</th></tr>
<tr><td><code>str.strip()</code> 吃掉 <code>\\t</code></td><td>用 Word COM 取段文本后 <code>.strip()</code> 会把制表符去掉，导致目录条目"结论\\t"被误判为正文标题"结论"，页码登记成目录所在页。判断是否为目录条目必须用<b>未 strip</b> 的原始文本。</td></tr>
<tr><td>页码独立成 run</td><td>目录条目可能拆成 <code>[标题][\\t31][30]</code> 三个 run，只改 tab run 会留下重复数字。必须"删除 tab run 之后所有 run + 页码写回 tab run"。</td></tr>
<tr><td>分页符藏在空行里</td><td>python-docx 读 <code>&lt;w:br w:type="page"/&gt;</code> 得到 <code>\\n</code> 而非 <code>\\x0c</code>，判断分页符要查 XML。</td></tr>
<tr><td>合并角注不能自动化</td><td>"同一观点"必须逐句人工判断；<code>[5]</code>（技术规程）与 <code>[15][16]</code>（地方标准）若各支撑不同分句，不能盲目合并。</td></tr>
<tr><td>Word COM 与 python-docx 索引不同</td><td><code>doc.Paragraphs</code> 含表格内段落，索引远大于 python-docx 的 <code>d.paragraphs</code>，且不可下标访问，需先缓存成 list。</td></tr>
</table></div>

</div></body></html>"""
out = os.path.join(BASE, "_格式核验报告_终版.html")
with io.open(out, "w", encoding="utf-8") as f:
    f.write(HTML)
print("已生成:", out)
