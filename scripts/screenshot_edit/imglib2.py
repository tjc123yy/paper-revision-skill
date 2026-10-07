# -*- coding: utf-8 -*-
"""学生H论文系统截图精修（依据 修改意见.txt 第4、5条）。

产出 media_new/*.png，随后由 apply_media.py 打回 docx 并同步图片显示尺寸。
"""
import os
import sys
import json
import re
from collections import Counter

import numpy as np
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import imglib as L

MEDIA = os.path.join(HERE, 'media')
OUT = os.path.join(HERE, 'media_new')
os.makedirs(OUT, exist_ok=True)

OCR = json.load(open(os.path.join(HERE, 'ocr.json'), encoding='utf-8'))

NEW_ACCOUNT = '2365000000000'
NEW_NAME = '学生H'
NEW_ADDR = '巴音郭楞蒙古自治州'
BRAND = '青丰禾润奶牛场智能养殖管理系统'
NAVY = (70, 98, 161)
BAND_H = 46


# ------------------------------------------------------------------ OCR 检索
def words(img):
    for ln in OCR[img]['lines']:
        for w in ln:
            yield ln, w


def ocr_boxes(img, pred):
    out = []
    for w in OCR[img]['lines']:
        for x in w:
            if pred(x['t']):
                out.append(x)
    return out


def word_at(img, x, y, tol=6):
    best = None
    for w in OCR[img]['lines']:
        for c in w:
            if abs(c['x'] - x) <= tol and abs(c['y'] - y) <= tol:
                best = c
    return best


def label_box(img, keys, ymin=0, ymax=100000):
    best = None
    for ln in OCR[img]['lines']:
        ws = ln
        for i in range(len(ws)):
            for j in range(i + 1, min(len(ws), i + 6) + 1):
                t = ''.join(w['t'] for w in ws[i:j]).replace(' ', '')
                if any(k in t for k in keys) and len(t) <= 6:
                    x0 = min(w['x'] for w in ws[i:j])
                    x1 = max(w['x'] + w['w'] for w in ws[i:j])
                    y0 = min(w['y'] for w in ws[i:j])
                    y1 = max(w['y'] + w['h'] for w in ws[i:j])
                    if not (ymin <= (y0 + y1) / 2.0 <= ymax):
                        continue
                    cand = (x0, y0, x1, y1, t)
                    if best is None or len(t) > len(best[4]):
                        best = cand
    return best


# ------------------------------------------------------------------ 绘制原语
def sample_bg(a, box, pad=8):
    x0, y0, x1, y1 = box
    h, w = a.shape[:2]
    rx0, ry0 = max(0, x0 - pad), max(0, y0 - pad)
    rx1, ry1 = min(w, x1 + pad + 1), min(h, y1 + pad + 1)
    region = a[ry0:ry1, rx0:rx1].copy()
    m = np.ones(region.shape[:2], dtype=bool)
    m[y0 - ry0:y1 - ry0 + 1, x0 - rx0:x1 - rx0 + 1] = False
    px = region[m].reshape(-1, 3)
    if len(px) == 0:
        px = region.reshape(-1, 3)
    return Counter(map(tuple, px)).most_common(1)[0][0]


def paint_text(a, box, text, bold=False, bg=None, color=None, align='center',
               pad_w=4, pad_h=3, max_w=None, thr=None, grow_x=0, grow_y=0):
    x0, y0, x1, y1 = box
    if bg is None:
        bg = sample_bg(a, box)
    if color is None:
        color = L.text_color(a, box)
    bw = x1 - x0 + 1
    bh = y1 - y0 + 1
    avail_w = max_w if max_w else bw
    size = L.fit_size(text, avail_w + pad_w, bh + pad_h, bold=bold)
    bb = L.ink_bbox(text, size, bold)
    iw, ih = bb[2] - bb[0] + 1, bb[3] - bb[1] + 1
    # 擦除范围要比文字 bbox 略大：深色阈值取到的 bbox 不含反锯齿边缘，
    # 只按 bbox 填充会在原文字上下留一圈"残影"。
    fx0, fy0 = x0 - grow_x, y0 - grow_y
    fx1, fy1 = x1 + grow_x, y1 + grow_y
    fx0 = max(0, fx0)
    fy0 = max(0, fy0)
    fx1 = min(a.shape[1] - 1, fx1)
    fy1 = min(a.shape[0] - 1, fy1)
    a[fy0:fy1 + 1, fx0:fx1 + 1] = bg
    bw = fx1 - fx0 + 1
    bh = fy1 - fy0 + 1
    x0, y0 = fx0, fy0
    canvas = Image.new('RGB', (iw + 40, ih + 40), (255, 255, 255))
    d = ImageDraw.Draw(canvas)
    d.text((20 - bb[0], 20 - bb[1]), text, font=L.font(size, bold), fill=tuple(int(v) for v in color))
    src = np.asarray(canvas).astype(np.float32)[20:20 + ih, 20:20 + iw]
    base = float(sum(int(v) for v in color)) / 3.0
    cov = (255.0 - src.min(axis=2)) / max(1.0, 255.0 - base)
    cov = np.clip(cov, 0.0, 1.0)
    cx = x0 + (bw - iw) // 2 if align == 'center' else x0
    cy = y0 + (bh - ih) // 2
    h, w = a.shape[:2]
    if cx < 0:
        cx = 0
    if cy < 0:
        cy = 0
    tw = min(iw, w - cx)
    th = min(ih, h - cy)
    cov = cov[:th, :tw]
    region = a[cy:cy + th, cx:cx + tw].astype(np.float32)
    out = region * (1.0 - cov[:, :, None]) + np.array([float(v) for v in color], np.float32) * cov[:, :, None]
    a[cy:cy + th, cx:cx + tw] = np.clip(out, 0, 255).astype(np.uint8)


def paint_lines(a, box, lines, color=None, thr=None, line_gap=2):
    """把一个多行文本块整体重绘（用于表格里换行的日期）。"""
    x0, y0, x1, y1 = box
    bg = sample_bg(a, box)
    if color is None:
        color = L.text_color(a, box)
    a[y0:y1 + 1, x0:x1 + 1] = bg
    n = len(lines)
    total = (y1 - y0 + 1) - line_gap * (n - 1)
    lh = total // n
    for i, ln in enumerate(lines):
        ly0 = y0 + i * (lh + line_gap)
        ly1 = ly0 + lh - 1
        paint_text(a, (x0, ly0, x1, ly1), ln, color=color, bg=bg, align='center',
                   pad_w=2, pad_h=1, thr=thr)


# ------------------------------------------------------------------ 版面分析
def table_columns(a, img, thr=190):
    vl = L.vertical_lines(a)
    top, bot = L.table_span(a, vl)
    out = {}
    for kind, keys in (('account', ['账号']), ('name', ['姓名']),
                       ('addr', ['位置']), ('date', ['时间']), ('date', ['日期'])):
        hb = label_box(img, keys, ymin=top - 6, ymax=top + 160)
        if not hb:
            continue
        band = L.column_band(vl, (hb[0] + hb[2]) / 2.0)
        if not band:
            continue
        cells = L.cells_in_band(a, band[0], band[1], top, bot, dark_thr=thr)
        out.setdefault(kind, (band, cells, hb[:4]))
    return vl, top, bot, out


def clean_cells(cells, band):
    bw = band[1] - band[0]
    out = []
    for c in cells:
        if (c[2] - c[0] + 1) > 0.92 * bw and (c[3] - c[1] + 1) < 12:
            continue
        if (c[3] - c[1] + 1) < 5:
            continue
        out.append(c)
    return out


def manual_column(a, x0, x1, thr=190, y0=0, y1=None):
    if y1 is None:
        y1 = a.shape[0] - 1
    cells = L.cells_in_band(a, x0, x1, y0, y1, dark_thr=thr)
    return (x0, x1), clean_cells(cells, (x0, x1))


# ------------------------------------------------------------------ 年份字形
def text_lines(mask, y0, y1, x0, x1, gap=3):
    sub = mask[y0:y1 + 1, x0:x1 + 1]
    rp = sub.sum(axis=1)
    return L.runs(list(rp > 0), gap=gap, minlen=3)


def glyph_groups(mask, x0, x1, y0, y1, thr=150):
    sub = (mask[y0:y1 + 1, x0:x1 + 1])
    cp = sub.sum(axis=0)
    groups = L.runs(list(cp > 0), gap=0, minlen=1)
    out = []
    for g in groups:
        bb = L.bbox_of(mask, x0 + g[0], x0 + g[1], y0, y1)
        if bb:
            out.append(bb)
    return out


def fix_year_in_cell(a, cell, verbose=False):
    """单元格内定位年份并只替换最后一位数字 3 -> 6。"""
    x0, y0, x1, y1 = cell
    m = L.dark_mask(a, 190)
    lines = text_lines(m, y0, y1, x0, x1)
    if not lines:
        return False
    ly0, ly1 = y0 + lines[0][0], y0 + lines[0][1]
    bb = L.bbox_of(m, x0, x1, ly0, ly1)
    if not bb:
        return False
    gs = glyph_groups(m, bb[0], bb[2], bb[1], bb[3])
    if verbose:
        print('      cell', cell, 'line1', bb, 'glyphs', len(gs), gs)
    if len(gs) < 9:
        return False
    gx0, gy0, gx1, gy1 = gs[3]
    paint_text(a, (gx0, gy0, gx1, gy1), '6',
               color=L.text_color(a, (bb[0], bb[1], bb[2], bb[3])),
               bg=sample_bg(a, (gx0 - 2, gy0 - 2, gx1 + 2, gy1 + 2), pad=3),
               pad_w=0, pad_h=2)
    return True


# ------------------------------------------------------------------ 品牌标识
def brand_band(a, text=BRAND, height=BAND_H):
    h, w = a.shape[:2]
    nb = Image.new('RGB', (w, h + height), tuple(NAVY))
    na = np.asarray(nb).copy()
    na[height:] = a
    # 左侧白色小圆角块 + 名称
    dr = ImageDraw.Draw(Image.fromarray(na))
    img = Image.fromarray(na)
    dr = ImageDraw.Draw(img)
    dr.rounded_rectangle([14, (height - 28) // 2, 42, (height - 28) // 2 + 28], radius=6, fill=(255, 255, 255))
    f = L.font(21, True)
    dr.text((30, height // 2), '青', font=f, fill=NAVY, anchor='mm')
    fb = L.font(21, False)
    dr.text((56, height // 2), text, font=fb, fill=(255, 255, 255), anchor='lm')
    dr.line([0, height - 1, w, height - 1], fill=(255, 255, 255), width=1)
    return np.asarray(img)


def brand_navbar(a, box, text=BRAND):
    x0, y0, x1, y1 = box
    img = Image.fromarray(a)
    dr = ImageDraw.Draw(img)
    dr.text((x0, (y0 + y1) // 2), text, font=L.font(20, False),
            fill=(255, 255, 255), anchor='lm')
    return np.asarray(img)


def brand_title(a, box, text=BRAND):
    img = Image.fromarray(a)
    dr = ImageDraw.Draw(img)
    a = np.asarray(img)
    paint_text(a, box, text, bold=True,
               color=L.text_color(a, box), bg=sample_bg(a, tuple(int(v) for v in box), pad=4),
               max_w=int((box[2] - box[0]) * 1.6))
    return a
