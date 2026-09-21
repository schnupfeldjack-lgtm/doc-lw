# -*- coding: utf-8 -*-
"""论文三配图：BIM协同设计"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from matplotlib import font_manager

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'figs_p3')
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


# 图3-1：BIM协同设计实施方案主要环节（6个环节纵向）
fig, ax = plt.subplots(figsize=(6.0, 7.5))
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis('off')

steps = [
    '参数化构件库构建\n（剪力墙/叠合板/楼梯/阳台板族库）',
    '多专业模型整合\n（建筑/结构/机电/装饰）',
    '碰撞检查与节点优化\n（结构与管线、设备预留洞口、节点可行性）',
    '构件深化设计\n（尺寸精度、预埋件位置、节点做法）',
    '设计-生产数据贯通\n（模型导出工厂生产管理系统）',
    '施工模拟与现场配合\n（拼装顺序、节点连接、净空控制）',
]

w = 6.5
h = 1.1
y_top = 8.2
for i, s in enumerate(steps):
    y = y_top - i * (h + 0.3)
    box(ax, (10 - w) / 2, y - h, w, h, s)
    if i < len(steps) - 1:
        arrow(ax, 10 / 2, y, 10 / 2, y - 0.3)

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig3_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

# 图4-1：BIM协同设计对结构设计质量的影响机理（四维度）
fig, ax = plt.subplots(figsize=(7.2, 5.2))
ax.set_xlim(0, 10)
ax.set_ylim(0, 8)
ax.axis('off')

box(ax, 3.4, 6.9, 3.2, 0.85, 'BIM协同设计\n对结构设计质量的影响', fc='#f2f2f2', fs=10.5)

dims = [
    ('设计协同', ['统一BIM模型下的\n多专业实时协同', '结构方案早期介入\n减少后期变更', '上下层与构件\n布置协同加强']),
    ('构件深化', ['预制构件尺寸\n精度提高', '预埋件位置\n准确性提高', '节点做法\n可行性提高']),
    ('管线综合', ['结构与机电管线\n碰撞识别', '预留洞口与配筋\n冲突识别', '净空优化\n提升居住品质']),
    ('设计变更', ['参数化变更自动\n传播减少不一致', '协同共享减少\n信息不对称', '三维检查减少\n二维盲区错误']),
]

xs = [0.1, 2.65, 5.2, 7.75]
bw = 2.15
for (title, items), x0 in zip(dims, xs):
    box(ax, x0, 5.15, bw, 0.85, title, fc='#f2f2f2', fs=10)
    arrow(ax, 5.0, 6.9, x0 + bw / 2, 6.0)
    for j, it in enumerate(items):
        y = 4.3 - j * 0.95
        box(ax, x0 + 0.1, y, bw - 0.2, 0.75, it, fs=9)

ax.text(5.0, 0.2, '整体效果：设计一次性成品率提高、设计错误率下降、设计周期可控',
        ha='center', fontsize=9.5,
        bbox=dict(facecolor='#fffacd', edgecolor='#888', boxstyle='round,pad=0.25'))

plt.tight_layout()
plt.savefig(os.path.join(OUT, 'fig4_1.png'), dpi=200, bbox_inches='tight', facecolor='white')
plt.close()

print('figs ->', OUT)
for f in os.listdir(OUT):
    print('  ', f, os.path.getsize(os.path.join(OUT, f)))