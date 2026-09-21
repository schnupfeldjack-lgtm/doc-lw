# -*- coding: utf-8 -*-
"""论文三：BIM协同设计（浦江镇S8-01保障房）"""
import os, shutil
from fill_docx import fill

ROOT = os.path.dirname(os.path.abspath(__file__))
PAPER = os.path.abspath(os.path.join(ROOT, '..', '..', 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'))
DOCX = 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例.docx'
TPL_SRC = os.path.abspath(os.path.join(ROOT, '..', '..', '炎黄职业技术学院毕业论文模板(1).docx'))

shutil.copyfile(TPL_SRC, os.path.join(PAPER, DOCX))

TITLE = 'BIM协同设计对建筑结构设计质量的影响研究——以装配式住宅项目为例'

MD_BODY = [
    os.path.join(ROOT, 'p3_part1.md'),
    os.path.join(ROOT, 'p3_part2.md'),
    os.path.join(ROOT, 'p3_part3.md'),
]
MD_META = os.path.join(ROOT, 'p3_meta.md')

FIG_DIR = os.path.join(ROOT, 'figs_p3')
FIGURES = {
    'fig3_1': os.path.join(FIG_DIR, 'fig3_1.png'),
    'fig4_1': os.path.join(FIG_DIR, 'fig4_1.png'),
}

fill(PAPER, DOCX, TITLE, MD_BODY, MD_META, FIGURES)
print('论文三生成：', os.path.join(PAPER, DOCX))