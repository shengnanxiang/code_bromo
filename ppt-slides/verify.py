# -*- coding: utf-8 -*-
"""验证 PPTX 文件内容：页数、中文文本、图片嵌入、页码徽章"""
from pptx import Presentation

PPTX = r'C:/Users/Nan/Documents/GitHub/code_bromo/Bromo产品定义介绍.pptx'
OUT = r'C:/Users/Nan/Documents/GitHub/code_bromo/ppt-slides/verify_result.txt'

prs = Presentation(PPTX)
lines = []
lines.append(f'Total slides: {len(prs.slides)}')
lines.append('=' * 60)

for i, slide in enumerate(prs.slides, 1):
    texts = []
    img_count = 0
    for shape in slide.shapes:
        if shape.shape_type == 13:  # PICTURE
            img_count += 1
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t:
                texts.append(t)
    first = texts[0][:45] if texts else '(no text)'
    lines.append(f'Slide {i:2d}: imgs={img_count}, texts={len(texts)} | {first}')

lines.append('=' * 60)
# 详细检查 slide 3 的中文内容
lines.append('--- Slide 3 full text ---')
for shape in prs.slides[2].shapes:
    if shape.has_text_frame:
        t = shape.text_frame.text.strip()
        if t:
            lines.append(f'  {t}')

lines.append('--- Slide 6 (comparison table) full text ---')
for shape in prs.slides[5].shapes:
    if shape.has_text_frame:
        t = shape.text_frame.text.strip()
        if t:
            lines.append(f'  {t}')

# 检查所有文本是否有乱码字符（常见乱码范围）
import re
lines.append('--- Mojibake check ---')
mojibake_pattern = re.compile(r'[\u4dc0-\u4dff\ue000-\uf8ff]')
bad_count = 0
all_text = []
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text
            all_text.append((i, t))

# 检查典型乱码序列（UTF-8被误读为GBK的特征）
gbk_mojibake = ['鏄', '浠', '涔', '鍥', '澶', '鏍', '浠', '寮', '鎷', '鏈', '虹', '鐨', '蹇', '璺', '鏁', '瓙', '鑳', '嵎']
found = False
for i, t in all_text:
    for ch in gbk_mojibake:
        if ch in t:
            lines.append(f'  WARNING: slide {i} contains possible mojibake char: {ch}')
            found = True
            bad_count += 1
            break
if not found:
    lines.append('  OK: No mojibake characters detected.')

# 检查占位符残留
lines.append('--- Placeholder check ---')
placeholder_found = False
for i, t in all_text:
    if '{' in t or 'TODO' in t or 'PLACEHOLDER' in t:
        lines.append(f'  WARNING: slide {i} may contain placeholder: {t[:50]}')
        placeholder_found = True
if not placeholder_found:
    lines.append('  OK: No placeholders detected.')

result = '\n'.join(lines)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(result)
print('Verification complete. Result written to', OUT)
