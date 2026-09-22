# -*- coding: utf-8 -*-
"""为再生混凝土论文重写版生成4幅原创技术图。"""
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p1_v2"
OUT.mkdir(parents=True, exist_ok=True)

# GitHub runner 安装 Noto/文泉等中文字体后自动选一个可用字体
available = {f.name for f in font_manager.fontManager.ttflist}
for cand in ["Noto Sans CJK SC", "Noto Sans CJK JP", "AR PL UKai CN", "SimHei", "Microsoft YaHei"]:
    if cand in available:
        plt.rcParams["font.sans-serif"] = [cand]
        break
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.size"] = 10

EDGE = "#333333"
FILL = "#f2f2f2"
FILL2 = "#ffffff"
MID = "#d9d9d9"

def box(ax, x, y, w, h, text, fc=FILL2, fs=9.5, lw=1.0):
    p = FancyBboxPatch((x,y),w,h,boxstyle="round,pad=0.02,rounding_size=0.05",
                       linewidth=lw,edgecolor=EDGE,facecolor=fc)
    ax.add_patch(p)
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",fontsize=fs)
    return p

def arrow(ax,x1,y1,x2,y2):
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),
                arrowprops=dict(arrowstyle="-|>",lw=1.0,color=EDGE,shrinkA=2,shrinkB=2))

def save(fig,name):
    fig.tight_layout()
    fig.savefig(OUT/name,dpi=220,bbox_inches="tight",facecolor="white")
    plt.close(fig)

# 图2-1：材料特征 -> 生产风险的传导
fig,ax=plt.subplots(figsize=(7.4,4.8))
ax.set_xlim(0,12);ax.set_ylim(0,7);ax.axis("off")
top=[
    (0.25,"旧砂浆附着\n孔隙率增加"),
    (2.65,"吸水率与\n批次离散增大"),
    (5.05,"有效用水与\n工作性波动"),
    (7.45,"早期强度与\n成型状态波动"),
    (9.85,"脱模、外观与\n尺寸风险"),
]
for x,t in top: box(ax,x,4.8,1.9,1.0,t,fc=FILL,fs=9.2)
for i in range(len(top)-1): arrow(ax,top[i][0]+1.9,5.3,top[i+1][0],5.3)
controls=[
    (0.45,"来源编码\n分级检验"),
    (3.15,"含水率实测\n动态修正"),
    (5.85,"经时工作性\n与早强验证"),
    (8.55,"首件预制\n脱模/外观检查"),
]
for x,t in controls: box(ax,x,1.9,2.5,0.9,t,fc=FILL2,fs=9)
for x,_ in controls:
    arrow(ax,x+1.25,2.8,x+1.25,4.8)
ax.text(6,6.45,"风险传导链",ha="center",va="center",fontsize=11,weight="bold")
ax.text(6,0.85,"控制原则：把材料波动转化为可测量、可修正、可追溯的生产参数",
        ha="center",fontsize=9.2)
save(fig,"fig2_1.png")

# 图3-1：官方事实链，不把案例冒充装配式
fig,ax=plt.subplots(figsize=(7.4,4.8))
ax.set_xlim(0,12);ax.set_ylim(0,7);ax.axis("off")
facts=[
    (0.4,4.55,2.45,1.15,"广西大学\n研究生公寓"),
    (3.15,4.55,2.45,1.15,"30层\n百米级高层"),
    (5.9,4.55,2.45,1.15,"再生骨料\n取代率30%"),
    (8.65,4.55,2.85,1.15,"再生混凝土\n10 679 m³"),
]
for x,y,w,h,t in facts: box(ax,x,y,w,h,t,fc=FILL,fs=9.8)
for i in range(3): arrow(ax,facts[i][0]+facts[i][2],5.12,facts[i+1][0],5.12)
box(ax,1.0,2.3,4.2,1.05,"可直接证明\n高层承重结构的规模化工程应用",fc=FILL2,fs=9.5)
box(ax,6.8,2.3,4.2,1.05,"不能直接证明\n预制生产、接缝和装配式节点性能",fc=FILL2,fs=9.5)
arrow(ax,5.2,2.82,6.8,2.82)
ax.text(6,6.35,"百米级再生混凝土工程的可核验事实链",ha="center",fontsize=11,weight="bold")
ax.text(6,0.95,"证据边界：案例是高层再生混凝土工程证据，不等同于装配式构件工程",
        ha="center",fontsize=9.2)
save(fig,"fig3_1.png")

# 图4-1：证据分层矩阵
fig,ax=plt.subplots(figsize=(7.4,5.4))
ax.set_xlim(0,12);ax.set_ylim(0,8.2);ax.axis("off")
cols=[0.4,3.2,6.15,9.1]
widths=[2.8,2.95,2.95,2.5]
headers=["评价对象","直接工程证据","装配式迁移证据","项目专项验证"]
for x,w,t in zip(cols,widths,headers):
    ax.add_patch(Rectangle((x,6.75),w,0.8,edgecolor=EDGE,facecolor=MID,linewidth=1))
    ax.text(x+w/2,7.15,t,ha="center",va="center",fontsize=9.3,weight="bold")
rows=[
    ("材料强度与总体施工","●","○","△"),
    ("预制生产稳定性","△","○","●"),
    ("框架/剪力墙构件","△","●","●"),
    ("叠合板等水平构件","△","●","●"),
    ("连接节点与接缝","—","●","●"),
    ("环境与经济性","△","○","●"),
]
y=5.95
for label,a,b,c in rows:
    vals=[label,a,b,c]
    for x,w,t in zip(cols,widths,vals):
        ax.add_patch(Rectangle((x,y),w,0.8,edgecolor=EDGE,facecolor="white",linewidth=0.85))
        ax.text(x+w/2,y+0.4,t,ha="center",va="center",fontsize=9.3)
    y-=0.8
ax.text(6,7.9,"再生混凝土向装配式住宅迁移的证据分层矩阵",
        ha="center",fontsize=11,weight="bold")
ax.text(0.65,0.7,"● 主要依据    ○ 辅助依据    △ 有限依据    — 不构成直接证据",
        ha="left",fontsize=9)
ax.text(6,0.25,"原则：不同证据只回答其能够支撑的问题，不用综合分数掩盖证据差异",
        ha="center",fontsize=9.2)
save(fig,"fig4_1.png")

# 图5-1：质量闸门
fig,ax=plt.subplots(figsize=(7.4,5.3))
ax.set_xlim(0,12);ax.set_ylim(0,8);ax.axis("off")
stages=[
    (0.35,5.25,2.05,1.15,"闸门1\n来源/批次准入"),
    (2.72,5.25,2.05,1.15,"闸门2\n含水率/试配"),
    (5.09,5.25,2.05,1.15,"闸门3\n首件预制"),
    (7.46,5.25,2.05,1.15,"闸门4\n构件/节点验证"),
    (9.83,5.25,1.82,1.15,"放行\n批量生产"),
]
for x,y,w,h,t in stages: box(ax,x,y,w,h,t,fc=FILL,fs=9.1)
for i in range(len(stages)-1):
    arrow(ax,stages[i][0]+stages[i][2],5.82,stages[i+1][0],5.82)
checks=[
    (0.55,"级配/吸水率\n含水率/来源"),
    (2.95,"工作性经时\n早期强度"),
    (5.3,"脱模/外观\n裂缝/尺寸"),
    (7.68,"承载/变形\n接缝可靠性"),
    (9.9,"批次追溯\n抽检/异常隔离"),
]
for x,t in checks: box(ax,x,2.85,1.65,1.05,t,fc=FILL2,fs=8.6)
for i,(x,_) in enumerate(checks):
    sx=stages[i][0]+stages[i][2]/2
    arrow(ax,sx,5.25,x+0.825,3.9)
# rejection loop
ax.annotate("异常退回前一环节复核",xy=(1.1,1.55),xytext=(10.8,1.55),
            ha="center",va="center",fontsize=9,
            arrowprops=dict(arrowstyle="-|>",lw=1,color=EDGE,
                            connectionstyle="arc3,rad=0.0"))
ax.text(6,7.15,"装配式再生混凝土的质量闸门与放行流程",
        ha="center",fontsize=11,weight="bold")
ax.text(6,0.65,"批量生产不是一次性审批，而是以批次数据持续证明生产状态受控",
        ha="center",fontsize=9.2)
save(fig,"fig5_1.png")

for p in sorted(OUT.glob("*.png")):
    print(p.name,p.stat().st_size)
