# -*- coding: utf-8 -*-
"""学生H论文系统截图精修主程序。产出 media_new/。"""
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
import imglib2 as L2

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
LOG = []


# --------------------------------------------------------------- 基础几何
def line_boxes(a, cell, gap=3, thr=190):
    x0, y0, x1, y1 = cell
    m = L.dark_mask(a, thr)
    sub = m[y0:y1 + 1, x0:x1 + 1]
    rp = sub.sum(axis=1)
    out = []
    for g in L.runs(list(rp >= 2), gap=gap, minlen=3):
        bb = L.bbox_of(m, x0, x1, y0 + g[0], y0 + g[1])
        if bb:
            out.append(bb)
    return out


def ocr_year_boxes(img):
    out = []
    for ln in OCR[img]['lines']:
        for w in ln:
            if w['t'] == '2023':
                out.append(w)
    return out


def ocr_month01_boxes(img):
    """仅取"紧跟在 2023 后面"的 01，避开时间里的 01。"""
    res = []
    for ln in OCR[img]['lines']:
        years = [w for w in ln if w['t'] == '2023']
        for w in ln:
            if w['t'] != '01':
                continue
            for y in years:
                if y['x'] + y['w'] - 2 <= w['x'] <= y['x'] + y['w'] + 50:
                    res.append(w)
                    break
    return res


def year_offset(a, cells, yboxes, thr=190):
    """用 OCR 年份框反推"年份相对单元格首行左端"的横向偏移。"""
    for b in yboxes:
        for c in cells:
            if c[0] <= b['x'] and b['x'] + b['w'] <= c[2] + 4 and c[1] - 4 <= b['y'] <= c[3] + 4:
                lines = line_boxes(a, c, thr=thr)
                if not lines:
                    continue
                l1 = lines[0]
                return (b['x'] - l1[0], b['x'] + b['w'] - l1[0])
    return None


def year_by_proportion(line, frac=0.45):
    """兜底：日期首行形如 2023-MM-DD(10 字符)，年份占行宽约 45%。"""
    x0, y0, x1, y1 = line
    w = x1 - x0 + 1
    return (x0, y0, x0 + int(round(frac * w)) - 1, y1)


def match_year_box(a, line, thr=190):
    """模板匹配定位首行开头的"2023"（无 OCR 兜底）。"""
    x0, y0, x1, y1 = line
    ink_h = y1 - y0 + 1
    best = None
    for k in (0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.45):
        size = max(8, int(ink_h * 1.35 * k))
        bb = L.ink_bbox('2023', size)
        tw, th = bb[2] - bb[0] + 1, bb[3] - bb[1] + 1
        if th > ink_h + 6 or tw > (x1 - x0) + 4:
            continue
        canvas = Image.new('L', (tw + 8, th + 8), 255)
        ImageDraw.Draw(canvas).text((4 - bb[0], 4 - bb[1]), '2023', font=L.font(size), fill=0)
        tmpl = np.asarray(canvas) < 200
        th2, tw2 = tmpl.shape
        m = L.dark_mask(a, thr)
        tn = tmpl.sum()
        for oy in (y0 - 2, y0 - 1, y0, y0 + 1):
            if oy < 0 or oy + th2 > a.shape[0]:
                continue
            for ox in range(x0 - 2, min(x1, x0 + int((x1 - x0) * 0.6)) + 1):
                if ox < 0 or ox + tw2 > a.shape[1]:
                    continue
                win = m[oy:oy + th2, ox:ox + tw2]
                score = (win & tmpl).sum() / float(tn)
                if best is None or score > best[0]:
                    best = (score, ox, oy, tw2, th2, size)
    if best is None or best[0] < 0.5:
        return None
    _, ox, oy, tw2, th2, size = best
    return (ox, oy, ox + tw2 - 1, oy + th2 - 1)


# --------------------------------------------------------------- 单项编辑
def set_year(a, box, thr=190):
    if box is None:
        return False
    box = tuple(int(round(v)) for v in box)
    L2.paint_text(a, box, '2026', color=L.text_color(a, box),
                  bg=L2.sample_bg(a, box, pad=3), pad_w=2, pad_h=3, thr=thr,
                  grow_x=1, grow_y=3)
    return True


def set_month(a, box, thr=190):
    if box is None:
        return False
    box = tuple(int(round(v)) for v in box)
    L2.paint_text(a, box, '08', color=L.text_color(a, box),
                  bg=L2.sample_bg(a, box, pad=3), pad_w=2, pad_h=3, thr=thr,
                  grow_x=1, grow_y=3)
    return True


def put(a, box, text, max_w=None, align='center', thr=190, grow_x=4, grow_y=6):
    if box is None:
        return False
    box = tuple(int(round(v)) for v in box)
    L2.paint_text(a, box, text, color=L.text_color(a, box),
                  bg=L2.sample_bg(a, box, pad=3), max_w=max_w, align=align, thr=thr,
                  grow_x=grow_x, grow_y=grow_y)
    return True


def value_right_of(a, lb, max_span=260, ypad=5, min_gap=6, thr=150):
    """表单：在标签右侧同一行带内找输入值（严格阈值，避开输入框边框）。"""
    x0, y0, x1, y1 = int(lb[0]), int(lb[1]), int(lb[2]), int(lb[3])
    ry0 = max(0, y0 - ypad)
    ry1 = min(a.shape[0] - 1, y1 + ypad)
    rx1 = min(a.shape[1] - 1, x1 + max_span)
    m = L.dark_mask(a, thr)
    bb = L.bbox_of(m, x1 + min_gap, rx1, ry0, ry1)
    return bb
