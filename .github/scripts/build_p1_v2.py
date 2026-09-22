# -*- coding: utf-8 -*-
"""从炎黄模板重新生成再生混凝土论文重写版。"""
from pathlib import Path
import shutil, sys

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / "新建文件夹/新建文件夹"
WORK = BASE / ".workbuddy-ai/work"
PAPER = BASE / "再生混凝土在装配式建筑中的应用评价——以住宅项目为例"
TPL = BASE / "炎黄职业技术学院毕业论文模板(1).docx"
OUT = PAPER / "再生混凝土在装配式建筑中的应用评价——以住宅项目为例_最终优化版.docx"

sys.path.insert(0, str(WORK))
from fill_docx import fill

def main():
    PAPER.mkdir(parents=True, exist_ok=True)
    if not TPL.exists():
        raise FileNotFoundError(TPL)
    shutil.copy2(TPL, OUT)
    figures = {
        "fig2_1": str(WORK / "figs_p1_v2/fig2_1.png"),
        "fig3_1": str(WORK / "figs_p1_v2/fig3_1.png"),
        "fig4_1": str(WORK / "figs_p1_v2/fig4_1.png"),
        "fig5_1": str(WORK / "figs_p1_v2/fig5_1.png"),
    }
    for p in figures.values():
        if not Path(p).exists():
            raise FileNotFoundError(p)
    fill(
        str(PAPER),
        OUT.name,
        "再生混凝土在装配式建筑中的应用评价——以住宅项目为例",
        [str(WORK/"p1_v2_part1.md"), str(WORK/"p1_v2_part2.md"), str(WORK/"p1_v2_part3.md")],
        str(WORK/"p1_v2_meta.md"),
        figures,
        target_chars_per_page=1200,
    )
    print("BUILT", OUT, OUT.stat().st_size)

if __name__ == "__main__":
    main()
