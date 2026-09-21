# -*- coding: utf-8 -*-
"""论文一：再生混凝土在装配式建筑中的应用评价"""
import os, shutil
from fill_docx import fill

ROOT = os.path.dirname(os.path.abspath(__file__))
PAPER = os.path.abspath(os.path.join(ROOT, '..', '..', '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'))
DOCX = '再生混凝土在装配式建筑中的应用评价——以住宅项目为例.docx'
TPL_SRC = os.path.abspath(os.path.join(ROOT, '..', '..', '炎黄职业技术学院毕业论文模板(1).docx'))

# 每次重新从原模板复制（保证母段落完整）
shutil.copyfile(TPL_SRC, os.path.join(PAPER, DOCX))

TITLE = '再生混凝土在装配式建筑中的应用评价——以住宅项目为例'

MD_BODY = [
    os.path.join(ROOT, 'p1_part1.md'),
    os.path.join(ROOT, 'p1_part2.md'),
    os.path.join(ROOT, 'p1_part3.md'),
]
MD_META = os.path.join(ROOT, 'p1_meta.md')

FIG_DIR = os.path.join(ROOT, 'figs_p1')
FIGURES = {
    'fig3_1': os.path.join(FIG_DIR, 'fig3_1.png'),
    'fig4_1': os.path.join(FIG_DIR, 'fig4_1.png'),
}

fill(PAPER, DOCX, TITLE, MD_BODY, MD_META, FIGURES)
print('论文一生成：', os.path.join(PAPER, DOCX))