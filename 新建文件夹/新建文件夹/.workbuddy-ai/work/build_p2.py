# -*- coding: utf-8 -*-
"""论文二：装配式施工质量管理（绵阳综合管廊）"""
import os, shutil
from fill_docx import fill

ROOT = os.path.dirname(os.path.abspath(__file__))
PAPER = os.path.abspath(os.path.join(ROOT, '..', '..', '装配式施工质量管理问题及优化研究——以市政项目为例'))
DOCX = '装配式施工质量管理问题及优化研究——以市政项目为例.docx'
TPL_SRC = os.path.abspath(os.path.join(ROOT, '..', '..', '炎黄职业技术学院毕业论文模板(1).docx'))

shutil.copyfile(TPL_SRC, os.path.join(PAPER, DOCX))

TITLE = '装配式施工质量管理问题及优化研究——以市政项目为例'

MD_BODY = [
    os.path.join(ROOT, 'p2_part1.md'),
    os.path.join(ROOT, 'p2_part2.md'),
    os.path.join(ROOT, 'p2_part3.md'),
]
MD_META = os.path.join(ROOT, 'p2_meta.md')

FIG_DIR = os.path.join(ROOT, 'figs_p2')
FIGURES = {
    'fig3_1': os.path.join(FIG_DIR, 'fig3_1.png'),
    'fig5_1': os.path.join(FIG_DIR, 'fig5_1.png'),
}

fill(PAPER, DOCX, TITLE, MD_BODY, MD_META, FIGURES)
print('论文二生成：', os.path.join(PAPER, DOCX))