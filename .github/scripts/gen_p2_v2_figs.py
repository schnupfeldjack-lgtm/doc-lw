# -*- coding: utf-8 -*-
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p2_v2"
OUT.mkdir(parents=True, exist_ok=True)

fonts = [f.name for f in font_manager.fontManager.ttflist]
for cand in ["Noto Sans CJK SC", "Noto Sans CJK JP", "AR PL UKai CN", "SimHei", "Microsoft YaHei"]:
    if cand in fonts:
        plt.rcParams["font.sans-serif"]=[cand]
        break
plt.rcParams["axes.unicode_minus"]=False

def box(ax,x,y,w,h,text,fs=10,lw=1.2,fc="white"):
    p=FancyBboxPatch((x,y),w,h,boxstyle="round,pad=0.02,rounding_size=0.06",
                     edgecolor="black",facecolor=fc,linewidth=lw)
    ax.add_patch(p)
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",fontsize=fs)

def arrow(ax,x1,y1,x2,y2):
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),
                arrowprops=dict(arrowstyle="-|>",linewidth=1.2,color="black",shrinkA=3,shrinkB=3))

# 图1-1：质量链—质量门—证据包
fig,ax=plt.subplots(figsize=(7.2,4.6))
ax.set_xlim(0,10);ax.set_ylim(0,7);ax.axis("off")
stages=["设计信息","工厂产品","运输交付","现场拼装","接口封闭","质量归档"]
xs=[0.2,1.85,3.5,5.15,6.8,8.45]
for i,(x,s) in enumerate(zip(xs,stages)):
    box(ax,x,5.35,1.35,0.75,s,9.5)
    if i<len(xs)-1: arrow(ax,x+1.35,5.72,xs[i+1],5.72)
ax.text(0.05,6.45,"质量链：质量状态随阶段传递",fontsize=10.5,weight="bold")
gates=["G0\n设计冻结","G1\n工厂放行","G2\n进场验收","G3\n拼装几何","G4\n隐蔽接口","G5\n数据闭环"]
for i,(x,g) in enumerate(zip(xs,gates)):
    box(ax,x,3.55,1.35,0.8,g,9)
    arrow(ax,x+0.68,5.35,x+0.68,4.35)
ax.text(0.05,4.6,"质量门：达不到条件，不进入下一阶段",fontsize=10.5,weight="bold")
evid=["版本/接口确认","尺寸/预埋/标识","到场状态/影像","测量记录/趋势","连接/防水记录","问题关闭/档案"]
for i,(x,e) in enumerate(zip(xs,evid)):
    box(ax,x,1.5,1.35,0.95,e,8.3)
    arrow(ax,x+0.68,3.55,x+0.68,2.45)
ax.text(0.05,2.8,"证据包：每次放行都有可追溯依据",fontsize=10.5,weight="bold")
ax.text(5,0.45,"核心逻辑：检查结果必须与“停止/放行”决策绑定",ha="center",fontsize=10)
plt.tight_layout();plt.savefig(OUT/"fig1_1.png",dpi=220,bbox_inches="tight",facecolor="white");plt.close()

# 图2-1：缺陷逃逸与放大
fig,ax=plt.subplots(figsize=(7.2,4.8))
ax.set_xlim(0,10);ax.set_ylim(0,7);ax.axis("off")
rows=[
 ("前序产生","模具/预埋偏差","运输碰损","基准偏差","基层或连接异常"),
 ("未被截断","出厂只查表面","进场只查合格证","单点合格但趋势失控","记录不完整仍封闭"),
 ("后序放大","安装不匹配","接缝修补","收口困难/返工","渗漏或无法追溯"),
]
ys=[5.4,3.5,1.6]
for ri,(lab,*items) in enumerate(rows):
    ax.text(0.15,ys[ri]+0.45,lab,fontsize=10.5,weight="bold",va="center")
    for j,it in enumerate(items):
        x=1.7+j*2.05
        box(ax,x,ys[ri],1.7,0.9,it,8.7)
        if ri<2: arrow(ax,x+0.85,ys[ri],x+0.85,ys[ri+1]+0.9)
ax.text(5,6.65,"缺陷不是“后序突然出现”，而是前序状态被带病放行",ha="center",fontsize=11,weight="bold")
plt.tight_layout();plt.savefig(OUT/"fig2_1.png",dpi=220,bbox_inches="tight",facecolor="white");plt.close()

# 图3-1：案例质量形成链
fig,ax=plt.subplots(figsize=(7.2,4.4))
ax.set_xlim(0,10);ax.set_ylim(0,6.5);ax.axis("off")
st=["标准化设计","构件生产","出厂交付","运输/堆放","吊装拼装","连接/防水","回填/归档"]
sub=[
 "构件编号\n接口冻结",
 "模具/钢筋\n预埋/养护",
 "尺寸/外观\n资料齐套",
 "保护状态\n身份一致",
 "轴线/标高\n接缝趋势",
 "隐蔽检查\n防水连续",
 "实体位置\n问题关闭",
]
xx=[0.1,1.5,2.9,4.3,5.7,7.1,8.5]
for i,x in enumerate(xx):
    box(ax,x,4.45,1.25,0.72,st[i],8.5)
    box(ax,x,2.75,1.25,0.9,sub[i],8.2)
    arrow(ax,x+0.625,4.45,x+0.625,3.65)
    if i<len(xx)-1:arrow(ax,x+1.25,4.81,xx[i+1],4.81)
ax.text(5,5.85,"装配式综合管廊质量形成链",ha="center",fontsize=11,weight="bold")
ax.text(5,1.55,"风险集中在阶段交接处：实体状态与责任同时发生转移",ha="center",fontsize=10)
ax.plot([0.4,9.6],[1.1,1.1],color="black",linewidth=1)
ax.text(5,0.55,"管理目标：前序问题前序关闭，禁止把“待确认”状态传给下一阶段",ha="center",fontsize=9.5)
plt.tight_layout();plt.savefig(OUT/"fig3_1.png",dpi=220,bbox_inches="tight",facecolor="white");plt.close()

# 图5-1：G0-G5质量门闭环
fig,ax=plt.subplots(figsize=(7.2,5.2))
ax.set_xlim(0,10);ax.set_ylim(0,8);ax.axis("off")
g=[
 ("G0 设计冻结","版本、接口、编号"),
 ("G1 工厂放行","本体、预埋、标识"),
 ("G2 进场验收","运输后状态"),
 ("G3 拼装几何","当前偏差+趋势"),
 ("G4 隐蔽接口","连接、防水、复验"),
 ("G5 数据闭环","实体—记录—整改"),
]
positions=[(0.3,5.7),(3.45,5.7),(6.6,5.7),(6.6,3.0),(3.45,3.0),(0.3,3.0)]
for (title,sub),(x,y) in zip(g,positions):
    box(ax,x,y,2.75,1.25,title+"\n"+sub,9.2,fc="#f7f7f7")
# arrows around loop
for i in range(len(positions)):
    x,y=positions[i]; nx,ny=positions[(i+1)%len(positions)]
    if i in [0,1]: arrow(ax,x+2.75,y+0.625,nx,ny+0.625)
    elif i==2: arrow(ax,x+1.375,y,nx+1.375,ny+1.25)
    elif i in [3,4]: arrow(ax,x,y+0.625,nx+2.75,ny+0.625)
    else: arrow(ax,x+1.375,y+1.25,positions[0][0]+1.375,positions[0][1])
box(ax,3.55,0.85,2.9,1.0,"异常闭环\n隔离→整改→复验→重新放行",9.5)
for x,y in positions:
    arrow(ax,x+1.375,y,x+1.375,1.85 if y<4 else 2.25)
ax.text(5,7.5,"六道质量门：用“停止条件”把全过程管理变成可执行动作",ha="center",fontsize=11,weight="bold")
plt.tight_layout();plt.savefig(OUT/"fig5_1.png",dpi=220,bbox_inches="tight",facecolor="white");plt.close()

for p in sorted(OUT.glob("*.png")):
    print(p.name,p.stat().st_size)
