# -*- coding: utf-8 -*-
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle, Polygon, Circle
from matplotlib import font_manager

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "新建文件夹/新建文件夹/.workbuddy-ai/work/figs_p2_polished"
OUT.mkdir(parents=True, exist_ok=True)

fonts = [f.name for f in font_manager.fontManager.ttflist]
for cand in ["Noto Sans CJK SC","Noto Sans CJK JP","Microsoft YaHei","SimHei","AR PL UKai CN"]:
    if cand in fonts:
        plt.rcParams["font.sans-serif"]=[cand]
        break
plt.rcParams["axes.unicode_minus"]=False

NAVY="#1F4E79"; BLUE="#5B9BD5"; LBLUE="#EAF2F8"
GREEN="#70AD47"; LGREEN="#EAF4E6"; ORANGE="#ED7D31"; LORANGE="#FCEFE6"
PURPLE="#8064A2"; LPURPLE="#F0ECF7"; RED="#C94C4C"; LRED="#FBE9E9"
TEAL="#2F8F9D"; LTEAL="#E6F4F6"; GRAY="#5B6573"; LGRAY="#F3F5F7"; GRID="#B9C3CE"

def setup(figsize=(7.4,4.7), xlim=(0,10), ylim=(0,6)):
    fig,ax=plt.subplots(figsize=figsize)
    ax.set_xlim(*xlim); ax.set_ylim(*ylim); ax.axis("off")
    fig.patch.set_facecolor("white"); ax.set_facecolor("white")
    return fig,ax

def round_box(ax,x,y,w,h,text,fc="white",ec=NAVY,fs=9.4,weight="normal",lw=1.25, radius=0.08, color="#1F2937"):
    p=FancyBboxPatch((x,y),w,h,boxstyle=f"round,pad=0.025,rounding_size={radius}",
                     facecolor=fc,edgecolor=ec,linewidth=lw)
    ax.add_patch(p)
    ax.text(x+w/2,y+h/2,text,ha="center",va="center",fontsize=fs,weight=weight,color=color,linespacing=1.25)
    return p

def arrow(ax,x1,y1,x2,y2,color=NAVY,lw=1.5,style="-|>",ls="-"):
    ax.annotate("",xy=(x2,y2),xytext=(x1,y1),
                arrowprops=dict(arrowstyle=style,linewidth=lw,color=color,shrinkA=2,shrinkB=2,linestyle=ls))

def pill(ax,x,y,w,h,text,fc,ec,fs=8.2,color="#243447"):
    round_box(ax,x,y,w,h,text,fc=fc,ec=ec,fs=fs,lw=1.0,radius=0.15,color=color)

def save(fig,name):
    fig.tight_layout(pad=0.45)
    fig.savefig(OUT/name,dpi=300,bbox_inches="tight",facecolor="white")
    plt.close(fig)

# 图1-1 质量链—质量门—证据包
fig,ax=setup((7.4,4.55),(0,10),(0,6))
ax.text(0.18,5.58,"质量链",fontsize=11,weight="bold",color=NAVY)
stages=["设计冻结","构件生产","运输交付","现场拼装","接口封闭","资料归档"]
sx=[0.7,2.2,3.7,5.2,6.7,8.2]
for i,(x,s) in enumerate(zip(sx,stages)):
    round_box(ax,x,4.75,1.15,0.62,s,fc=LBLUE,ec=BLUE,fs=9.1,weight="bold",color=NAVY)
    if i<len(sx)-1: arrow(ax,x+1.15,5.06,sx[i+1],5.06,color=BLUE,lw=1.8)
ax.text(0.18,3.61,"质量门",fontsize=11,weight="bold",color=ORANGE)
gates=["G0\n设计冻结门","G1\n工厂放行门","G2\n进场验收门","G3\n拼装几何门","G4\n隐蔽接口门","G5\n数据闭环门"]
for x,g in zip(sx,gates):
    round_box(ax,x,3.00,1.15,0.78,g,fc=LORANGE,ec=ORANGE,fs=8.7,weight="bold",color="#8A3B12")
    arrow(ax,x+0.575,4.75,x+0.575,3.78,color=ORANGE,lw=1.1,ls="--")
ax.text(0.18,1.65,"证据包",fontsize=11,weight="bold",color=GREEN)
evid=["冻结清单","首件/出厂记录","到场验收单","测量与复测","隐蔽/旁站记录","闭环/归档清单"]
for x,e in zip(sx,evid):
    round_box(ax,x,1.05,1.15,0.72,e,fc=LGREEN,ec=GREEN,fs=7.8,color="#235A2B")
    arrow(ax,x+0.575,3.00,x+0.575,1.77,color=GREEN,lw=1.0,ls="--")
ax.add_patch(FancyBboxPatch((0.65,0.22),8.72,0.48,boxstyle="round,pad=0.02,rounding_size=0.08",
                            facecolor=LGRAY,edgecolor=GRID,linewidth=0.9))
ax.text(5.01,0.46,"核心：检查必须绑定“停止 / 放行”决策，状态未确认不得跨阶段流转",
        ha="center",va="center",fontsize=8.8,color=GRAY)
save(fig,"fig1_1.png")

# 图2-1 质量形成机理
fig,ax=setup((7.4,4.7),(0,10),(0,6.35))
ax.text(5,6.03,"多因素耦合贯穿全过程，质量在跨阶段交接中形成",ha="center",fontsize=10.8,weight="bold",color=NAVY)
factors=[("人员能力",BLUE,LBLUE),("材料构件",GREEN,LGREEN),("设备机具",ORANGE,LORANGE),
         ("工法工序",PURPLE,LPURPLE),("环境条件",RED,LRED),("信息协同",TEAL,LTEAL)]
fx=[0.3,1.93,3.56,5.19,6.82,8.45]
for x,(t,c,fc) in zip(fx,factors):
    round_box(ax,x,5.02,1.28,0.64,t,fc=fc,ec=c,fs=8.6,weight="bold",color=c)
    arrow(ax,x+0.64,5.02,x+0.64,4.36,color=c,lw=1.0)
ax.add_patch(FancyBboxPatch((0.32,2.45),9.33,1.66,boxstyle="round,pad=0.02,rounding_size=0.08",
                            facecolor="#FAFCFE",edgecolor=GRID,linewidth=1.0))
proc=["设计信息","工厂制造","运输堆放","吊装拼接","连接防水","成品保护","交付运维"]
px=[0.5,1.83,3.16,4.49,5.82,7.15,8.48]
pcs=[BLUE,GREEN,ORANGE,PURPLE,RED,TEAL,NAVY]
for i,(x,t,c) in enumerate(zip(px,proc,pcs)):
    pill(ax,x,3.28,1.02,0.48,t,fc="white",ec=c,fs=7.7,color=c)
    if i<len(px)-1: arrow(ax,x+1.02,3.52,px[i+1],3.52,color=GRID,lw=1.4)
details=["方案深化","生产检验","保护交付","定位校正","连接密封","保护监测","验收归档"]
for x,t in zip(px,details):
    ax.text(x+0.51,2.84,t,ha="center",va="center",fontsize=7.4,color=GRAY)
ax.text(0.38,4.33,"质量形成链",fontsize=9.1,weight="bold",color=NAVY)
results=[("几何精度","尺寸、轴线、标高与累计偏差",BLUE,LBLUE),
         ("连接可靠性","连接构造、密实性与整体性",ORANGE,LORANGE),
         ("耐久与防水","接缝连续、抗渗与长期服役",GREEN,LGREEN)]
rx=[0.75,3.65,6.55]
for x,(t,sub,c,fc) in zip(rx,results):
    round_box(ax,x,0.72,2.45,0.92,t+"\n"+sub,fc=fc,ec=c,fs=8.2,weight="bold",color=c)
    arrow(ax,x+1.225,2.45,x+1.225,1.64,color=c,lw=1.15)
ax.text(5,0.22,"质量不是末端抽检的单点结果，而是设计—生产—物流—拼装—接口—信息全过程共同形成",
        ha="center",fontsize=8.2,color=GRAY)
save(fig,"fig2_1.png")

# 图3-1 案例工程质量形成链
fig,ax=setup((7.4,4.65),(0,10),(0,6.2))
ax.text(5,5.88,"案例工程：六种质量状态与阶段交接口",ha="center",fontsize=10.8,weight="bold",color=NAVY)
states=["可生产","可出厂","可进场","可拼装","可隐蔽","可归档"]
subs=["设计已冻结","本体+资料可交付","运输后状态未变","基准/支撑满足","连接/防水已确认","实体与记录已闭环"]
xv=[0.5,2.05,3.6,5.15,6.7,8.25]
colors=[BLUE,GREEN,ORANGE,PURPLE,TEAL,NAVY]
for i,(x,s,sub,c) in enumerate(zip(xv,states,subs,colors)):
    round_box(ax,x,3.85,1.2,0.75,s,fc="white",ec=c,fs=9.0,weight="bold",color=c,lw=1.5)
    ax.text(x+0.6,3.48,sub,ha="center",fontsize=7.3,color=GRAY)
    if i<len(xv)-1:
        arrow(ax,x+1.2,4.23,xv[i+1],4.23,color=GRID,lw=1.5)
        # gate diamond
        cx=(x+1.2+xv[i+1])/2
        poly=Polygon([[cx,4.63],[cx+0.18,4.23],[cx,3.83],[cx-0.18,4.23]],closed=True,facecolor=LORANGE,edgecolor=ORANGE,lw=1)
        ax.add_patch(poly); ax.text(cx,4.23,f"G{i}",ha="center",va="center",fontsize=7.6,weight="bold",color="#8A3B12")
risks=["版本/接口未定","批次性几何偏差","运输碰损/身份错配","累计偏差","隐蔽缺陷","资料断链"]
evid=["冻结清单","首件/出厂证据","到场状态影像","连续测量记录","隐蔽验收记录","闭环档案"]
for x,r,e,c in zip(xv,risks,evid,colors):
    pill(ax,x,2.25,1.2,0.46,r,fc=LRED,ec=RED,fs=7.1,color="#8C2D2D")
    arrow(ax,x+0.6,3.45,x+0.6,2.71,color=RED,lw=0.9,ls="--")
    pill(ax,x,1.15,1.2,0.46,e,fc=LGREEN,ec=GREEN,fs=7.1,color="#235A2B")
    arrow(ax,x+0.6,2.25,x+0.6,1.61,color=GREEN,lw=0.9,ls="--")
ax.text(0.25,2.48,"主要风险",fontsize=8.5,weight="bold",color=RED)
ax.text(0.25,1.38,"放行证据",fontsize=8.5,weight="bold",color=GREEN)
ax.add_patch(FancyBboxPatch((1.0,0.28),8.0,0.48,boxstyle="round,pad=0.02,rounding_size=0.08",
                            facecolor=LBLUE,edgecolor=BLUE,linewidth=0.9))
ax.text(5,0.52,"管理目标：前序异常必须在责任交接前关闭，禁止把“待确认”状态传给下一阶段",
        ha="center",va="center",fontsize=8.2,color=NAVY)
save(fig,"fig3_1.png")

# 图4-1 缺陷跨阶段逃逸
fig,ax=setup((7.4,4.85),(0,10),(0,6.55))
ax.text(5,6.2,"缺陷产生 → 质量门失效 → 跨阶段转移 → 后序放大",ha="center",fontsize=10.8,weight="bold",color=NAVY)
headers=[("问题源头",BLUE,LBLUE),("失效的质量门",ORANGE,LORANGE),("传播机制",PURPLE,LPURPLE),("后序表现",GREEN,LGREEN)]
hx=[0.25,2.78,5.31,7.84]
for x,(t,c,fc) in zip(hx,headers):
    round_box(ax,x,5.45,1.9,0.5,t,fc=fc,ec=c,fs=8.8,weight="bold",color=c)
rows=[
("模具/预埋系统偏差","G1 未识别批次异常","批量复制","拼装不匹配/返工"),
("运输碰损未记录","G2 只核证书不核状态","状态变化被忽略","连接面/止水构造受损"),
("安装小偏差连续累积","G3 只看单点不看趋势","累计放大","收口困难/接缝失衡"),
("隐蔽施工证据不足","G4/G5 仍允许封闭","缺陷掩盖+资料断链","渗漏/追责困难"),
]
ys=[4.45,3.3,2.15,1.0]
for ri,row in enumerate(rows):
    cols=[BLUE,ORANGE,PURPLE,GREEN]; fcs=[LBLUE,LORANGE,LPURPLE,LGREEN]
    for ci,(x,txt,c,fc) in enumerate(zip(hx,row,cols,fcs)):
        round_box(ax,x,ys[ri],1.9,0.74,txt,fc=fc,ec=c,fs=7.6,color="#243447")
        if ci<3: arrow(ax,x+1.9,ys[ri]+0.37,hx[ci+1],ys[ri]+0.37,color=GRID,lw=1.3)
ax.text(5,0.28,"关键点：问题并非“后序突然出现”，而是前序异常没有被停止条件截断",
        ha="center",fontsize=8.4,color=GRAY)
save(fig,"fig4_1.png")

# 图5-1 G0-G5六道质量门
fig,ax=setup((7.4,5.0),(0,10),(0,6.75))
ax.text(5,6.43,"G0—G5 六道质量门及证据闭环",ha="center",fontsize=10.8,weight="bold",color=NAVY)
gates=[
("G0","设计冻结门","图纸/接口/编号","设计单位",BLUE,LBLUE),
("G1","工厂放行门","首件/尺寸/预埋","预制厂",GREEN,LGREEN),
("G2","进场验收门","状态/编码/资料","施工总包",ORANGE,LORANGE),
("G3","拼装几何门","轴线/标高/趋势","施工总包",PURPLE,LPURPLE),
("G4","隐蔽接口门","连接/防水/复验","施工+监理",TEAL,LTEAL),
("G5","数据闭环门","归档/追溯/销项","项目各方",NAVY,LBLUE),
]
gx=[0.25,1.88,3.51,5.14,6.77,8.4]
for i,(x,(code,name,check,role,c,fc)) in enumerate(zip(gx,gates)):
    round_box(ax,x,4.62,1.35,0.72,code+"\n"+name,fc=fc,ec=c,fs=8.3,weight="bold",color=c,lw=1.4)
    pill(ax,x,3.65,1.35,0.52,check,fc="white",ec=c,fs=7.1,color=GRAY)
    pill(ax,x,2.75,1.35,0.46,"主控："+role,fc=LGRAY,ec=GRID,fs=7.0,color=GRAY)
    # decision
    poly=Polygon([[x+0.675,2.3],[x+0.97,1.86],[x+0.675,1.42],[x+0.38,1.86]],
                 closed=True,facecolor="white",edgecolor=c,lw=1.2)
    ax.add_patch(poly);ax.text(x+0.675,1.86,"满足\n放行条件？",ha="center",va="center",fontsize=6.9,color=c,weight="bold")
    if i<len(gx)-1:
        arrow(ax,x+0.97,1.86,gx[i+1]+0.38,1.86,color=GREEN,lw=1.35)
        ax.text(x+1.14,2.02,"是",fontsize=6.8,color=GREEN,weight="bold")
    arrow(ax,x+0.675,1.42,x+0.675,0.82,color=RED,lw=1.0)
    round_box(ax,x+0.25,0.32,0.85,0.42,"整改 / 复核",fc=LRED,ec=RED,fs=6.8,color=RED,lw=0.9)
    arrow(ax,x+0.25,0.53,x-0.05 if x>0.3 else x+0.1,3.9,color=RED,lw=0.75,style="-",ls="--")
# phase arrows
for i in range(len(gx)-1):
    arrow(ax,gx[i]+1.35,4.98,gx[i+1],4.98,color=GRID,lw=1.2)
ax.add_patch(FancyBboxPatch((1.05,5.65),7.9,0.46,boxstyle="round,pad=0.02,rounding_size=0.08",
                            facecolor=LGRAY,edgecolor=GRID,linewidth=0.8))
ax.text(5,5.88,"原则：达不到条件就停止流转；异常隔离—整改—复验—重新申请放行",
        ha="center",va="center",fontsize=8.2,color=GRAY)
save(fig,"fig5_1.png")

for p in sorted(OUT.glob("*.png")):
    print("POLISHED_FIG",p.name,p.stat().st_size)
