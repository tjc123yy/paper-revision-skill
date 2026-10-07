# -*- coding: utf-8 -*-
"""学生H 论文系统截图批量精修。

处理内容（依据 修改意见.txt 第4、5条）：
  1) 系统界面加"青丰禾润奶牛场智能养殖管理系统"标识（顶部蓝条 / 标题替换）
  2) 日期 2023 -> 2026，并把超出半年的 01 月改成 08 月
  3) 人名（真实姓名已隐去）-> 学生H
  4) 用户账号 ming/mlng -> 学号 2365000000000
  5) 地址 沧州 / 河北保定 / 河北沧州 -> 巴音郭楞蒙古自治州

实现方式：不改原截图结构，只在原图上做"局部擦除 + 同字体同色重绘"，逐像素级。
"""
import os
import sys
import json
import re
from collections import Counter

import numpy as np
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
MEDIA = os.path.join(HERE, 'media')
OUT = os.path.join(HERE, 'media_new')
os.makedirs(OUT, exist_ok=True)

FONT_REG = r'C:\Windows\Fonts\msyh.ttc'
FONT_BOLD = r'C:\Windows\Fonts\msyhbd.ttc'

NEW_ACCOUNT = '2365000000000'
NEW_NAME = '学生H'
NEW_ADDR = '巴音郭楞蒙古自治州'
NEW_YEAR = '2026'
BRAND = '青丰禾润奶牛场智能养殖管理系统'

_font_cache = {}


def font(size, bold=False):
    key = (int(size), bold)
    if key not in _font_cache:
        _font_cache[key] = ImageFont.truetype(FONT_BOLD if bold else FONT_REG, int(size))
    return _font_cache[key]


# ---------------------------------------------------------------- basic utils
def load(name):
    return np.asarray(Image.open(os.path.join(MEDIA, name)).convert('RGB')).astype(np.uint8)


def save(img, name):
    Image.fromarray(img).save(os.path.join(OUT, name))


def dark_mask(a, thr=190):
    return a.max(axis=2) < thr


def nonwhite_mask(a):
    return (a[:, :, 0] < 250) | (a[:, :, 1] < 250) | (a[:, :, 2] < 250)


def borderish_mask(a):
    """浅灰/浅蓝的表格边框线：近中性色且明显不是白。"""
    r, g, b = a[:, :, 0].astype(np.int16), a[:, :, 1].astype(np.int16), a[:, :, 2].astype(np.int16)
    mx = np.maximum(np.maximum(r, g), b)
    mn = np.minimum(np.minimum(r, g), b)
    return (mx < 250) & (mx > 95) & ((mx - mn) < 45)


def runs(flags, gap=2, minlen=1):
    out = []
    cur = []
    for i, f in enumerate(flags):
        if f:
            if cur and i - cur[-1] > gap:
                if len(cur) >= minlen:
                    out.append((cur[0], cur[-1]))
                cur = []
            cur.append(i)
    if len(cur) >= minlen:
        out.append((cur[0], cur[-1]))
    return out


def bbox_of(mask, x0, x1, y0, y1):
    sub = mask[y0:y1 + 1, x0:x1 + 1]
    ys, xs = np.nonzero(sub)
    if len(ys) == 0:
        return None
    return (int(xs.min()) + x0, int(ys.min()) + y0, int(xs.max()) + x0, int(ys.max()) + y0)


def ink_bbox(text, size, bold=False):
    f = font(size, bold)
    tmp = Image.new('L', (int(size * len(text) * 2 + 40), int(size * 3) + 40), 255)
    d = ImageDraw.Draw(tmp)
    d.text((20, 20), text, font=f, fill=0)
    a = np.asarray(tmp)
    m = a < 190
    ys, xs = np.nonzero(m)
    if len(ys) == 0:
        return None
    return (xs.min() - 20, ys.min() - 20, xs.max() - 20, ys.max() - 20)


def fit_size(text, max_w, max_h, bold=False, cap=48):
    """返回使 ink 尺寸不超过 (max_w, max_h) 且尽量接近的最大字号。"""
    ref = 100
    bb = ink_bbox(text, ref, bold)
    if not bb:
        return 12
    iw = bb[2] - bb[0] + 1
    ih = bb[3] - bb[1] + 1
    s = min(max_w / float(iw), max_h / float(ih)) * ref
    return max(7.0, min(cap, s))


def text_color(a, box):
    x0, y0, x1, y1 = box
    sub = a[y0:y1 + 1, x0:x1 + 1].reshape(-1, 3)
    lum = sub.astype(np.int16).sum(axis=1)
    idx = np.argsort(lum)[:max(1, len(lum) // 20)]
    return tuple(int(v) for v in sub[idx].mean(axis=0))


def bg_color(a, x0, x1, y0, y1, box):
    """取区域内除了文字框以外的众数颜色作为背景。"""
    sub = a[y0:y1 + 1, x0:x1 + 1].copy()
    sx0, sy0 = max(0, box[0] - x0 - 2), max(0, box[1] - y0 - 2)
    sx1, sy1 = min(sub.shape[1], box[2] - x0 + 3), min(sub.shape[0], box[3] - y0 + 3)
    m = np.ones(sub.shape[:2], dtype=bool)
    m[sy0:sy1, sx0:sx1] = False
    px = sub[m].reshape(-1, 3)
    if len(px) == 0:
        px = sub.reshape(-1, 3)
    cnt = Counter(map(tuple, px))
    return cnt.most_common(1)[0][0]


def draw_text(a, box, text, color, bold=False, align='center', pad=2):
    x0, y0, x1, y1 = box
    w = x1 - x0 + 1
    h = y1 - y0 + 1
    size = fit_size(text, w + pad, h + pad, bold=bold)
    f = font(size, bold)
    bb = ink_bbox(text, size, bold)
    iw = bb[2] - bb[0] + 1
    ih = bb[3] - bb[1] + 1
    canvas = Image.new('RGB', (w + 40, h + 40), (255, 255, 255))
    d = ImageDraw.Draw(canvas)
    d.text((20 - bb[0], 20 - bb[1]), text, font=f, fill=tuple(int(v) for v in color))
    ca = np.asarray(canvas)
    if align == 'center':
        ox = x0 + (w - iw) // 2
    else:
        ox = x0
    oy = y0 + (h - ih) // 2
    src = ca[20 - bb[1]:20 - bb[1] + ih, 20 - bb[0]:20 - bb[0] + iw]
    # 只把"笔画"贴上去，白底不覆盖（保护原有背景/条纹）
    srcmask = src.max(axis=2) < 250
    dst = a[oy:oy + ih, ox:ox + iw]
    dst[srcmask] = src[srcmask]
    a[oy:oy + ih, ox:ox + iw] = dst


# ---------------------------------------------------------------- table model
def vertical_lines(a):
    h, w, _ = a.shape
    bm = borderish_mask(a)
    frac = bm.sum(axis=0) / float(h)
    cols = [x for x in range(w) if frac[x] > 0.55]
    groups = runs([x in set(cols) for x in range(w)], gap=1, minlen=1)
    out = []
    for g in groups:
        if g[1] - g[0] <= 4:
            out.append((g[0] + g[1]) // 2)
        # 宽组是图片列，不是边框，丢弃
    return out


def table_span(a, vlines):
    bm = borderish_mask(a)
    h, w, _ = a.shape
    tops, bots = [], []
    for x in vlines:
        col = bm[:, x]
        ys = np.nonzero(col)[0]
        if len(ys) > h * 0.3:
            tops.append(ys.min())
            bots.append(ys.max())
    if not tops:
        return 0, h - 1
    return int(np.median(tops)), int(np.median(bots))


def cells_in_band(a, x0, x1, y0, y1, dark_thr=190, line_merge=30, min_h=4, inset=4, min_row_dark=2):
    x0 = x0 + inset
    x1 = x1 - inset
    m = dark_mask(a, dark_thr)
    sub = m[y0:y1 + 1, x0:x1 + 1]
    rowdark = sub.sum(axis=1)
    flags = list(rowdark >= min_row_dark)
    lines = runs(flags, gap=2, minlen=min_h)
    if not lines:
        return []
    merged = []
    for ln in lines:
        if merged and ln[0] - merged[-1][1] <= line_merge:
            merged[-1] = (merged[-1][0], ln[1])
        else:
            merged.append(list(ln))
    out = []
    for c in merged:
        bb = bbox_of(m, x0, x1, y0 + c[0], y0 + c[1])
        if bb:
            out.append(bb)
    return out


def column_band(vlines, xc):
    for i in range(len(vlines) - 1):
        if vlines[i] < xc < vlines[i + 1]:
            return vlines[i], vlines[i + 1]
    return None
