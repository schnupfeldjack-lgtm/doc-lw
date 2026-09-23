# -*- coding: utf-8 -*-
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p2_mono"
OUT.mkdir(parents=True, exist_ok=True)

fonts=[f.name for f in font_manager.fontManager.ttflist]
for cand in ["Noto Sans CJK SC","Noto Sans CJK JP","Microsoft YaHei","SimHei","AR PL UKai CN"]:
    if cand in fonts:
        plt.rcParams["font.sans-serif"]=[cand]; break
plt.rcParams["axes.unicode_minus"]=False

BLACK="#111111"; DARK="#333333"; MID="#777777"; GRID="#A0A0A0"; LIGHT="#F5F5F5"; WHITE="#FFFFFF"

def setup(figsize, xlim, ylim):
    fig,ax=plt.subplots(figsize=figsize)
    fig.patch.set_facecolor("white"); ax.set_facecolor("white")
    ax.set_xlim(*xlim); ax.set_ylim(*ylim); ax.axis("off")
    return fig,ax

def box(ax,x,y,w,h,text,fs=8.8,bold=False,fc=WHITE,ec=BLACK,lw=1.0,rad=0.06):
    p=FancyBboxPatch((x,y),w,h,boxstyle=f"round,pad=0.015,rounding_size={rad}",
                     facecolor=fc,edgecolor=ec,linewidth=lw)
    ax.add_patch(p)
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",fontsize=fs,
            weight="bold" if bold else "normal",color=BLACK,linespacing=1.2)
    return p

def arrow(ax,x1,y1,x2,y2,ls="-",lw=1.0):
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),
                arrowprops=dict(arrowstyle="-|>",color=BLACK,linewidth=lw,linestyle=ls,
                                shrinkA=2,shrinkB=2))

def save(fig,name):
    fig.tight_layout(pad=0.25)
    fig.savefig(OUT/name,dpi=320,bbox_inches="tight",facecolor="white",pad_inches=0.04)
    plt.close(fig)

# 图1-1：紧凑横向质量链—质量门—证据包
fig,ax=setup((7.2,3.15),(0,10),(0,4.2))
sx=[0.35,1.95,3.55,5.15,6.75,8.35]
stages=["设计冻结","构件生产","运输交付","现场拼装","接口封闭","资料归档"]
for i,(x,s) in enumerate(zip(sx,stages)):
    box(ax,x,3.22,1.25,0.52,s,fs=8.2,bold=True,lw=1.05)
    if i<5: arrow(ax,x+1.25,3.48,sx[i+1],3.48,lw=1.1)
gates=["G0","G1","G2","G3","G4","G5"]
for x,g in zip(sx,gates):
    poly=Polygon([[x+0.625,2.88],[x+0.78,2.58],[x+0.625,2.28],[x+0.47,2.58]],
                 closed=True,facecolor=WHITE,edgecolor=BLACK,lw=1.0)
    ax.add_patch(poly);ax.text(x+0.625,2.58,g,ha="center",va="center",fontsize=7.2,weight="bold")
    arrow(ax,x+0.625,3.22,x+0.625,2.88,lw=0.85)
evid=["冻结清单","首件/出厂记录","到场状态记录","测量/复测记录","隐蔽验收记录","闭环归档"]
for x,e in zip(sx,evid):
    box(ax,x,1.37,1.25,0.5,e,fs=7.0,fc=LIGHT,ec=DARK,lw=0.9)
    arrow(ax,x+0.625,2.28,x+0.625,1.87,ls="--",lw=0.8)
ax.text(0.08,3.48,"质量链",fontsize=8.7,weight="bold",va="center")
ax.text(0.08,2.58,"质量门",fontsize=8.7,weight="bold",va="center")
ax.text(0.08,1.62,"证据包",fontsize=8.7,weight="bold",va="center")
box(ax,1.0,0.32,8.0,0.48,"核心：未达到放行条件不得跨阶段流转；异常必须隔离、整改、复验并留痕",
    fs=7.5,fc=LIGHT,ec=GRID,lw=0.8)
save(fig,"fig1_1.png")

# 图2-1：质量形成机理，紧凑黑白
fig,ax=setup((7.2,3.35),(0,10),(0,4.6))
factors=["人员能力","材料构件","设备机具","工法工序","环境条件","信息协同"]
fx=[0.25,1.87,3.49,5.11,6.73,8.35]
for x,t in zip(fx,factors):
    box(ax,x,3.82,1.35,0.5,t,fs=7.9,bold=True,fc=LIGHT,ec=DARK,lw=0.9)
    arrow(ax,x+0.675,3.82,x+0.675,3.45,lw=0.75)
proc=["设计","生产","运输","拼装","连接/防水","保护","归档"]
px=[0.35,1.75,3.15,4.55,5.95,7.35,8.75]
for i,(x,t) in enumerate(zip(px,proc)):
    box(ax,x,2.75,0.95,0.48,t,fs=7.5,bold=True)
    if i<6:arrow(ax,x+0.95,2.99,px[i+1],2.99,lw=1.0)
sub=["信息冻结","构件检验","状态保护","几何校正","隐蔽确认","成品监测","资料闭环"]
for x,t in zip(px,sub):
    ax.text(x+0.475,2.46,t,ha="center",fontsize=6.7,color=DARK)
res=[("几何精度","尺寸/轴线/标高"),("连接可靠","连接/密实/整体"),("耐久防水","接缝/抗渗/服役")]
rx=[0.75,3.75,6.75]
for x,(a,b) in zip(rx,res):
    box(ax,x,0.85,2.5,0.72,a+"\n"+b,fs=7.7,bold=True,fc=LIGHT,ec=DARK,lw=0.9)
    arrow(ax,x+1.25,2.30,x+1.25,1.57,lw=0.8)
ax.text(5,0.25,"质量形成于全过程耦合，而非末端一次检验。",ha="center",fontsize=7.5,color=DARK)
save(fig,"fig2_1.png")

# 图3-1：案例六种状态与交接口，不在图内放标题，压缩空白
fig,ax=setup((7.2,3.45),(0,10),(0,4.7))
states=["可生产","可出厂","可进场","可拼装","可隐蔽","可归档"]
subs=["设计已冻结","本体+资料可交付","运输后状态未变","基准/支撑满足","连接/防水已确认","实体与记录已闭环"]
risks=["设计未定/接口未定","批次性几何偏差","运输碰损/身份错配","累计偏差","隐蔽缺陷","资料断链"]
evid=["冻结清单","首件/出厂证据","到场状态影像","连续测量记录","隐蔽验收记录","闭环档案"]
xv=[0.55,2.05,3.55,5.05,6.55,8.05]
for i,(x,s,sub) in enumerate(zip(xv,states,subs)):
    box(ax,x,3.72,1.22,0.54,s,fs=8.0,bold=True)
    ax.text(x+0.61,3.46,sub,ha="center",fontsize=6.5,color=DARK)
    if i<5:
        cx=(x+1.22+xv[i+1])/2
        poly=Polygon([[cx,4.18],[cx+0.15,3.99],[cx,3.80],[cx-0.15,3.99]],
                     closed=True,facecolor=WHITE,edgecolor=BLACK,lw=0.9)
        ax.add_patch(poly); ax.text(cx,3.99,f"G{i}",ha="center",va="center",fontsize=6.4,weight="bold")
        arrow(ax,x+1.22,3.99,cx-0.15,3.99,lw=0.8);arrow(ax,cx+0.15,3.99,xv[i+1],3.99,lw=0.8)
for x,r,e in zip(xv,risks,evid):
    box(ax,x,2.16,1.22,0.52,r,fs=6.6,fc=WHITE,ec=DARK,lw=0.85)
    arrow(ax,x+0.61,3.38,x+0.61,2.68,ls="--",lw=0.75)
    box(ax,x,0.92,1.22,0.48,e,fs=6.7,fc=LIGHT,ec=DARK,lw=0.85)
    arrow(ax,x+0.61,2.16,x+0.61,1.40,ls="--",lw=0.75)
ax.text(0.08,2.42,"主要风险",fontsize=8.0,weight="bold")
ax.text(0.08,1.16,"放行证据",fontsize=8.0,weight="bold")
box(ax,1.15,0.16,7.75,0.42,"管理目标：前序异常必须在责任交接前关闭，禁止把“待确认”状态传给下一阶段",
    fs=7.0,fc=LIGHT,ec=GRID,lw=0.75)
save(fig,"fig3_1.png")

# 图4-1：缺陷传播路径，黑白表意
fig,ax=setup((7.2,3.55),(0,10),(0,4.85))
headers=["问题源头","质量门失效","传播机制","后序表现"]
hx=[0.35,2.85,5.35,7.85]
for x,t in zip(hx,headers):
    box(ax,x,4.18,1.8,0.45,t,fs=8.0,bold=True,fc=LIGHT,ec=DARK,lw=0.9)
rows=[
("模具/预埋系统偏差","G1未识别批次异常","批量复制","拼装不匹配/返工"),
("运输碰损未记录","G2只核证书不核状态","状态变化被忽略","连接面/止水构造受损"),
("安装小偏差连续累积","G3只看单点不看趋势","累计放大","收口困难/接缝失衡"),
("隐蔽施工证据不足","G4/G5仍允许封闭","缺陷掩盖+资料断链","渗漏/追责困难")]
ys=[3.32,2.42,1.52,0.62]
for y,row in zip(ys,rows):
    for ci,(x,txt) in enumerate(zip(hx,row)):
        box(ax,x,y,1.8,0.55,txt,fs=6.7,fc=WHITE if ci%2==0 else LIGHT,ec=DARK,lw=0.75)
        if ci<3:arrow(ax,x+1.8,y+0.275,hx[ci+1],y+0.275,lw=0.85)
save(fig,"fig4_1.png")

# 图5-1：G0-G5质量门闭环，横向整齐、无斜飞线
fig,ax=setup((7.2,4.0),(0,10),(0,5.4))
box(ax,1.2,4.68,7.6,0.42,"原则：达不到条件就停止流转；异常隔离—整改—复验—重新申请放行",
    fs=7.3,fc=LIGHT,ec=GRID,lw=0.8)
gates=[
("G0","设计冻结门","图纸/接口/编号","主控：设计单位"),
("G1","工厂放行门","首件/尺寸/预埋","主控：预制厂"),
("G2","进场验收门","状态/编码/资料","主控：施工总包"),
("G3","拼装几何门","轴线/标高/趋势","主控：施工总包"),
("G4","隐蔽接口门","连接/防水/复验","主控：施工+监理"),
("G5","数据闭环门","归档/追溯/销项","主控：项目各方")]
gx=[0.25,1.87,3.49,5.11,6.73,8.35]
for i,(x,(code,name,check,role)) in enumerate(zip(gx,gates)):
    box(ax,x,3.88,1.35,0.55,code+"\n"+name,fs=7.5,bold=True,fc=WHITE,ec=BLACK,lw=0.95)
    if i<5:arrow(ax,x+1.35,4.155,gx[i+1],4.155,lw=0.95)
    box(ax,x,3.05,1.35,0.48,check,fs=6.5,fc=WHITE,ec=DARK,lw=0.75)
    arrow(ax,x+0.675,3.88,x+0.675,3.53,lw=0.75)
    box(ax,x,2.28,1.35,0.46,role,fs=6.25,fc=LIGHT,ec=GRID,lw=0.7)
    arrow(ax,x+0.675,3.05,x+0.675,2.74,lw=0.75)
    poly=Polygon([[x+0.675,2.00],[x+0.93,1.63],[x+0.675,1.26],[x+0.42,1.63]],
                 closed=True,facecolor=WHITE,edgecolor=BLACK,lw=0.9)
    ax.add_patch(poly);ax.text(x+0.675,1.63,"满足\n放行条件?",ha="center",va="center",fontsize=5.9)
    arrow(ax,x+0.675,2.28,x+0.675,2.00,lw=0.75)
    if i<5:
        arrow(ax,x+0.93,1.63,gx[i+1]+0.42,1.63,lw=0.85)
        ax.text(x+1.18,1.77,"是",fontsize=5.8,weight="bold")
    box(ax,x+0.16,0.45,1.03,0.42,"整改 / 复核",fs=6.3,fc=WHITE,ec=DARK,lw=0.75)
    arrow(ax,x+0.675,1.26,x+0.675,0.87,lw=0.75)
save(fig,"fig5_1.png")

for p in sorted(OUT.glob("*.png")):
    print("MONO_FIG",p.name,p.stat().st_size)
