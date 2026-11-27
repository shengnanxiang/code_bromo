# -*- coding: utf-8 -*-
"""定稿验证：修订项落实情况 + 文件完整性"""
from pptx import Presentation

PPTX = r'C:/Users/Nan/Documents/GitHub/code_bromo/Bromo产品定义介绍.pptx'
OUT = r'C:/Users/Nan/Documents/GitHub/code_bromo/ppt-slides/final_check.txt'

prs = Presentation(PPTX)
lines = []
lines.append(f'Total slides: {len(prs.slides)}')
lines.append('=' * 60)

# 收集每页全部文本
def slide_texts(idx):
    texts = []
    for shape in prs.slides[idx].shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t:
                texts.append(t)
    return texts

# --- 修订项 1: P10 图注（索引 9） ---
t10 = slide_texts(9)
ok1 = any('取景器内对焦确认灯及声音反馈' in t for t in t10)
lines.append(f'[修订1] P10 图注含"及声音反馈": {"PASS" if ok1 else "FAIL"}')

# --- 修订项 2: P13 图注（索引 12） ---
t13 = slide_texts(12)
ok2 = any('弹出后蓄电及完成指示' in t for t in t13)
ok2b = any('独立 2×AA 供电' in t for t in t13)
lines.append(f'[修订2] P13 图注"弹出后蓄电及完成指示": {"PASS" if ok2 else "FAIL"}')
lines.append(f'[修订2b] P13 第三图图注"独立 2×AA 供电": {"PASS" if ok2b else "FAIL"}')

# 检查 P13 图片坐标（第三张图应不再与价值锚点重叠）
from pptx.util import Emu
p13_pics = []
for shape in prs.slides[12].shapes:
    if shape.shape_type == 13:
        p13_pics.append((shape.left, shape.top, shape.width, shape.height))
p13_pics.sort(key=lambda p: p[1])
if len(p13_pics) >= 3:
    bottom_pic_bottom = Emu(p13_pics[-1][1] + p13_pics[-1][3]).inches
    lines.append(f'[修订2c] P13 最下图片底边: {bottom_pic_bottom:.2f}in (应 < 4.9): {"PASS" if bottom_pic_bottom < 4.9 else "FAIL"}')

# --- 修订项 3: P17 参数表（索引 16） ---
t17 = slide_texts(16)
ok3 = any('尺寸参考：Konica C35' in t for t in t17) and any('设计参考：Leica M' in t for t in t17)
lines.append(f'[修订3] P17 "设计参考/尺寸参考/布局参考"三列对仗: {"PASS" if ok3 else "FAIL"}')

# --- 修订项 4: P11 高度标注（索引 10） ---
t11 = slide_texts(10)
ok4 = any('高位 3.5–4mm' in t for t in t11) and any('沉入约 1mm' in t for t in t11)
lines.append(f'[修订4] P11 高度标注: {"PASS" if ok4 else "FAIL"}')

# --- 全局完整性 ---
lines.append('=' * 60)
img_total = 0
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.shape_type == 13:
            img_total += 1
lines.append(f'Images embedded total: {img_total} (expect 15)')

# 占位符与乱码复查
import re
placeholder_words = ['TODO', 'PLACEHOLDER', '{name}', '{title}']
ph_found = False
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text
            for w in placeholder_words:
                if w in t:
                    lines.append(f'  WARNING: slide {i} placeholder: {t[:40]}')
                    ph_found = True
if not ph_found:
    lines.append('Placeholders: none (PASS)')

# 每页文本量复核
lines.append('=' * 60)
for i, slide in enumerate(prs.slides, 1):
    texts = []
    for shape in slide.shapes:
        if shape.shape_type == 13:
            pass
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t:
                texts.append(t)
    first = texts[0][:30] if texts else '(no text)'
    lines.append(f'Slide {i:2d}: {len(texts):2d} texts | {first}')

result = '\n'.join(lines)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(result)
print('Final check written to', OUT)
