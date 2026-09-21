# -*- coding: utf-8 -*-
"""把图/表编号的分隔符改为 en dash（–），匹配模板"图3–5"写法"""
import os

p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fill_docx.py')
s = open(p, encoding='utf-8').read()

EN = '\u2013'  # en dash
old_fig = "f'图{chapter}-{figure_seq}'"
new_fig = "f'图{chapter}%s{figure_seq}'" % EN
old_tbl = "f'表{chapter}-1'"
new_tbl = "f'表{chapter}%s1'" % EN

n = 0
if old_fig in s:
    s = s.replace(old_fig, new_fig)
    n += 1
if old_tbl in s:
    s = s.replace(old_tbl, new_tbl)
    n += 1
open(p, 'w', encoding='utf-8').write(s)
print('图/表编号分隔符修正处数:', n)
