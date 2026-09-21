# -*- coding: utf-8 -*-
"""论文二配图：绵阳综合管廊施工质量控制"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib import font_manager

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'figs_p2')
os.makedirs(OUT, exist_ok=True)

names = set(f.name for f in font_manager.fontManager.ttflist)
for cand in ['SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi', 'FangSong']:
    if cand in names:
        plt.rcParams['font.sans-serif'] = [cand]
        break
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 10.5


def box(ax, x, y, w, h, text, fc='#ffffff', fs=10):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.012',
                                edgecolor='#000000', facecolor=fc, linewidth=1.1))
    ax.text(x + w / 2, y + h / 2, text, ha='center', va='center', fontsize=fs, zorder=3)


def arrow(ax, x1, y1, x2, y2):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='-|>', color='#000000', linewidth=1.1,
                                shrinkA=0, shrinkB=0, mutation_scale=12))


# 图3-1：施工质量控制主要环节
fig, ax = plt.subplots(figsize=(7.2, 5.2))
ax.set_xlim(0, 10)
ax.set_ylim(0, 7.5)
ax.axis('off')

steps_left = [
    '构件出厂检验\n（尺寸、预埋件、外观）',
    '进场验收与批次\n追溯建档',
    '节段运输与现场\n二次保护',
    '节段吊装就位\n（粗定位→精定位）',
    '三维扫描比对\n（动态修正）',
]
steps_right = [
    '节段接缝防水\n（卷材铺设、自愈合）',
    '灌浆套筒连接\n（密实度抽检）',
    '多专业穿插\n（工序协调、成品保护）',
    '关键工序举牌\n验收与影像留痕',
    '质量分析与\n改进闭环',
]

w, h = 4.0, 0.95
x1, x2 = 0.25, 5.75
ys = [6.35, 5.15, 3.95, 2.75, 1.55]

for i, s in enumerate(steps_left):
    box(ax, x1, ys[i], w, h, s)
for i, s in enumerate(steps_right):
    box(ax, x2, ys[i], w, h, s)

for i in range(4):
    arrow(ax, x1 + w / 2, ys[i], x1 + w / 2, ys[i + 1] + h)
    arrow(ax, x2 + w / 2, ys[i] + h, x2 + w / 2, ys[i + 1])
arrow(ax, x1 + w, ys[4] + h / 2, x2, ys[4] + h / 2)

ax.text(x1 + w / 2, 7.3, '构件质量与现场拼装阶段', ha='center', fontsize=10)
ax.text(x2 + w / 2, 7.3, '节点连接与质量改进阶段', ha='center', fontsize=10)

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig3_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

# 图5-1：优化对策框架
fig, ax = plt.subplots(figsize=(7.2, 5.0))
ax.set_xlim(0, 10)
ax.set_ylim(0, 7.5)
ax.axis('off')

box(ax, 3.5, 6.4, 3.0, 0.85, '装配式综合管廊施工\n质量管理优化对策', fc='#f2f2f2', fs=10.5)

dims = [
    ('组织管理优化', ['构件批次追溯制度', '驻厂监理制度', '工序交接与举牌验收']),
    ('工艺技术优化', ['BIM+三维扫描动态\n修正工艺', '预铺反粘+自愈合\n防水卷材复合工艺', '构件拆分与吊装\n工艺优化']),
    ('质量控制手段优化', ['数字化质量管理\n平台', '关键工序三检制\n与举牌验收', '质量分析与改进\n常态化机制']),
]

xs = [0.15, 3.45, 6.75]
bw = 3.1
for (title, items), x0 in zip(dims, xs):
    box(ax, x0, 4.7, bw, 0.85, title, fc='#f2f2f2', fs=10)
    arrow(ax, 5.0, 6.4, x0 + bw / 2, 5.55)
    for j, it in enumerate(items):
        y = 3.8 - j * 1.0
        box(ax, x0 + 0.25, y, bw - 0.5, 0.75, it, fs=9.5)
        arrow(ax, x0 + bw / 2, 4.7, x0 + bw / 2, y + 0.75)

ax.text(5.0, 0.15, '效果：预埋件安装合格率100%，结构质量零缺陷',
        ha='center', fontsize=9.5, color='#000000',
        bbox=dict(facecolor='#fffacd', edgecolor='#888', boxstyle='round,pad=0.25'))

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig5_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

print('figs ->', OUT)
for f in os.listdir(OUT):
    print('  ', f, os.path.getsize(os.path.join(OUT, f)))