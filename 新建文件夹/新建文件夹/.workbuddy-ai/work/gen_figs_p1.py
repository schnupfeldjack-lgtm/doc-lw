# -*- coding: utf-8 -*-
"""生成论文一配图（原创绘制）"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib import font_manager

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'figs_p1')
os.makedirs(OUT, exist_ok=True)

names = set(f.name for f in font_manager.fontManager.ttflist)
for cand in ['SimHei', 'Microsoft YaHei', 'SimSun', 'KaiTi', 'FangSong', 'STZhongsong']:
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


# ============ 图3-1 全过程质量控制要点 ============
fig, ax = plt.subplots(figsize=(7.2, 5.4))
ax.set_xlim(0, 10)
ax.set_ylim(0, 7.5)
ax.axis('off')

steps_left = [
    '废弃混凝土回收\n(来源甄别、分类堆放)',
    '破碎与筛分\n(除杂、粒径控制)',
    '分级检验\n(按 GB/T 25177 分级)',
    '预湿预处理\n(控制含水率波动)',
    '配合比设计\n(水胶比、掺合料)',
]
steps_right = [
    '工厂集中拌合\n(计量精度控制)',
    '出厂与运输\n(坍落度逐车检测)',
    '浇筑与振捣\n(防漏振、防过振)',
    '保湿养护\n(时间适当延长)',
    '实体质量检测\n(强度、尺寸偏差)',
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

# 跨列连接
arrow(ax, x1 + w, ys[4] + h / 2, x2, ys[4] + h / 2)

ax.text(x1 + w / 2, 7.3, '再生骨料加工与预处理阶段', ha='center', fontsize=10)
ax.text(x2 + w / 2, 7.3, '混凝土生产与施工阶段', ha='center', fontsize=10)

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig3_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

# ============ 图4-1 评价指标体系 ============
fig, ax = plt.subplots(figsize=(7.2, 5.0))
ax.set_xlim(0, 10)
ax.set_ylim(0, 8)
ax.axis('off')

box(ax, 3.4, 6.9, 3.2, 0.85, '再生混凝土在装配式\n住宅中的应用评价', fc='#f2f2f2', fs=10.5)

dims = [
    ('技术性能\n（能不能用）', ['强度性能', '变形性能', '耐久性能', '结构性能']),
    ('施工适应性\n（好不好用）', ['拌合物工作性', '泵送与浇筑', '养护要求', '构件生产效率']),
    ('经济与环境效益\n（值不值得用）', ['材料成本', '固废消纳量', '碳排放降低量', '综合经济性']),
]

xs = [0.15, 3.45, 6.75]
bw = 3.1
for (title, items), x0 in zip(dims, xs):
    box(ax, x0, 5.15, bw, 0.95, title, fc='#f2f2f2', fs=10)
    arrow(ax, 5.0, 6.9, x0 + bw / 2, 6.15)
    for j, it in enumerate(items):
        y = 4.15 - j * 0.9
        box(ax, x0 + 0.25, y, bw - 0.5, 0.68, it, fs=9.5)
        arrow(ax, x0 + bw / 2, 5.15, x0 + bw / 2, y + 0.68)

ax.text(5.0, 0.35, '注：技术性能是前提，施工适应性是关键，经济与环境效益是推广动力',
        ha='center', fontsize=9)

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig4_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

print('figs ->', OUT)
for f in os.listdir(OUT):
    print('  ', f, os.path.getsize(os.path.join(OUT, f)))
