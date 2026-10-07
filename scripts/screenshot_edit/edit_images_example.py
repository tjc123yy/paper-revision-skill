# -*- coding: utf-8 -*-
"""主驱动：按图逐个执行编辑。产出 media_new/。"""
import os
import sys
import json

import numpy as np
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import imglib as L
import imglib2 as L2
import edit_core as E

OUT = E.OUT
THR = 190

# (kind, x0, x1, skip_first)  kind: a=账号 n=姓名 ad=地址
MAN_COLS = {
    'image23.png': [('a', 982, 1087, False, 0, None), ('n', 1087, 1193, False, 0, None)],
    'image25.png': [('a', 710, 866, True, 200, 346), ('n', 888, 1040, True, 200, 346),
                    ('ad', 500, 684, True, 200, 346)],
    'image40.png': [('a', 920, 1050, True, 235, 399), ('n', 1050, 1200, True, 235, 399)],
    'image46.png': [('a', 176, 330, True, 200, 438), ('n', 330, 540, True, 200, 438)],
}
AUTO_COLS = {
    'image16.png': ['a', 'n', 'd'], 'image17.png': ['a', 'n', 'd'],
    'image18.png': ['a', 'n', 'd'], 'image20.png': ['a', 'n', 'd'],
    'image21.png': ['a', 'n', 'd'], 'image24.png': ['a', 'n', 'ad'],
    'image26.png': ['a', 'n', 'ad'], 'image28.png': ['a', 'n', 'ad'],
    'image29.png': ['a', 'n', 'd'], 'image31.png': ['a', 'n', 'd'],
    'image32.png': ['a', 'n', 'd'], 'image34.png': ['a', 'n', 'd'],
    'image35.png': ['a', 'n', 'd'], 'image39.png': ['a', 'n', 'd'],
    'image45.png': ['a', 'n'],
}
# 日期列需要手工给列带（OCR 未给出表头）
MAN_DATE_COLS = {
    'image23.png': (876, 982),
}
# 需要额外修月份的行（0 基，表头之后）：image32 最后一行为 2023-01-27
MONTH_FIX = {'image32.png': [3]}
PER_IMAGE_THR = {'image22.png': 110}
FORM = {
    'image15.png': [('账号', 'a'), ('姓名', 'n')],
    'image19.png': [('账号', 'a'), ('姓名', 'n')],
    'image27.png': [('位置', 'ad'), ('账号', 'a'), ('姓名', 'n')],
    'image30.png': [('账号', 'a'), ('姓名', 'n')],
    'image33.png': [('账号', 'a'), ('姓名', 'n')],
    'image36.png': [('账号', 'a'), ('姓名', 'n')],
    'image41.png': [('账号', 'a'), ('姓名', 'n')],
}
ROW_FILTER = {'image45.png': [0]}
AUTO_AVAIL = {'a': None, 'n': None, 'ad': 174}
FORM_AVAIL = {'a': 250, 'n': 120, 'ad': 200}
TXT = {'a': E.NEW_ACCOUNT, 'n': E.NEW_NAME, 'ad': E.NEW_ADDR}
KEYMAP = {'account': 'a', 'name': 'n', 'addr': 'ad', 'date': 'd'}
HEADER_KEYS = {'a': ['账号'], 'n': ['姓名'], 'ad': ['位置']}
BRAND_NAVBAR = {'image19.png', 'image36.png'}


def apply_common(a, img, thr, log):
    n_y = n_m = 0
    for w in E.ocr_year_boxes(img):
        box = (int(w['x']), int(w['y']), int(w['x'] + w['w'] - 1), int(w['y'] + w['h'] - 1))
        if E.set_year(a, box, thr):
            n_y += 1
    for w in E.ocr_month01_boxes(img):
        box = (int(w['x']), int(w['y']), int(w['x'] + w['w'] - 1), int(w['y'] + w['h'] - 1))
        if E.set_month(a, box, thr):
            n_m += 1
    log.append('  common: year=%d month=%d' % (n_y, n_m))


def apply_col(a, cells, kind, header=None, rows=None, avail_w=None, skip_first=False, thr=THR, log=None):
    cells = list(cells)
    if header is not None:
        cells = [c for c in cells if not overlaps(c, header)]
    elif skip_first:
        cells = cells[1:]
    if rows is not None:
        cells = [c for i, c in enumerate(cells) if i in rows]
    w = avail_w or (cells[0][2] - cells[0][0] + 1 if cells else 100)
    n = 0
    for c in cells:
        if E.put(a, c, TXT[kind], max_w=w, thr=thr):
            n += 1
    if log:
        log.append('  col %s: %d cells (avail=%s)' % (kind, n, w))
    return n


def overlaps(c, hb, pad=6):
    return not (c[2] < hb[0] - pad or c[0] > hb[2] + pad or c[3] < hb[1] - pad or c[1] > hb[3] + pad)


def apply_date_col(a, img, band, cells, thr, log, header=None, month_rows=None):
    yboxes = [w for w in E.ocr_year_boxes(img) if band[0] - 8 <= w['x'] <= band[1] + 8]
    off = E.year_offset(a, cells, yboxes, thr) if yboxes else None
    done = 0
    data = []
    for c in cells:
        if header is not None and overlaps(c, header):
            continue
        data.append(c)
        has = any(c[1] - 8 <= w['y'] <= c[3] + 8 and c[0] - 6 <= w['x'] <= c[2] + 6 for w in yboxes)
        if has:
            continue
        lines = E.line_boxes(a, c, thr)
        if not lines:
            continue
        l1 = lines[0]
        box = None
        if off:
            box = (l1[0] + off[0], l1[1], l1[0] + off[1], l1[3])
        else:
            box = E.match_year_box(a, l1, thr)
            if box is None:
                box = E.year_by_proportion(l1)
        if E.set_year(a, box, thr):
            done += 1
    # 月份兜底（OCR 未识别出 01 时）
    if month_rows:
        for idx in month_rows:
            if idx >= len(data):
                continue
            c = data[idx]
            lines = E.line_boxes(a, c, thr)
            if not lines:
                continue
            l1 = lines[0]
            if off:
                yx0, yx1 = l1[0] + off[0], l1[0] + off[1]
            else:
                yx0, yx1 = E.year_by_proportion(l1)[0], E.year_by_proportion(l1)[2]
            cw = (yx1 - yx0 + 1) / 4.0
            mx0 = int(round(yx1 + 1 + 0.45 * cw))
            mx1 = int(round(mx0 + 2 * cw))
            if E.set_month(a, (mx0, l1[1], mx1, l1[3]), thr):
                log.append('  month fallback row %d box=(%d,%d,%d,%d)' % (idx, mx0, l1[1], mx1, l1[3]))
    log.append('  date col: offset=%s extra_year=%d cells=%d' % (off, done, len(data)))


def process(img, log):
    a = L.load(img).copy()
    thr = PER_IMAGE_THR.get(img, THR)
    log.append('=== %s %dx%d thr=%d' % (img, a.shape[1], a.shape[0], thr))

    apply_common(a, img, thr, log)

    for keys, kind in FORM.get(img, []):
        lb = L2.label_box(img, [keys])
        if not lb:
            log.append('  form label %s NOT FOUND' % keys)
            continue
        vb = E.value_right_of(a, lb, thr=min(150, thr))
        if not vb:
            log.append('  form value %s NOT FOUND' % keys)
            continue
        E.put(a, vb, TXT[kind], max_w=FORM_AVAIL[kind], align='left', thr=min(150, thr))
        log.append('  form %s -> %s' % (keys, vb))

    if img in AUTO_COLS:
        vl, top, bot, cols = L2.table_columns(a, img, thr)
        for kind in AUTO_COLS[img]:
            key = {v: k for k, v in KEYMAP.items()}[kind]
            if key not in cols:
                log.append('  auto col %s NOT FOUND' % kind)
                continue
            band, cells, hbox = cols[key]
            cells = L2.clean_cells(cells, band)
            if kind == 'd':
                apply_date_col(a, img, band, cells, thr, log, header=hbox,
                               month_rows=MONTH_FIX.get(img))
            else:
                av = AUTO_AVAIL[kind]
                if av is None:
                    av = band[1] - band[0] - 12
                apply_col(a, cells, kind, header=hbox, avail_w=av, thr=thr, log=log)

    if img in MAN_DATE_COLS:
        x0, x1 = MAN_DATE_COLS[img]
        band, cells = L2.manual_column(a, x0, x1, thr, 0, None)
        cells = L2.clean_cells(cells, band)
        apply_date_col(a, img, band, cells, thr, log)

    for kind, x0, x1, sk, my0, my1 in MAN_COLS.get(img, []):
        band, cells = L2.manual_column(a, x0, x1, thr, my0, my1)
        cells = L2.clean_cells(cells, band)
        rows = ROW_FILTER.get(img)
        apply_col(a, cells, kind, rows=rows, avail_w=x1 - x0 - 12, skip_first=sk,
                  thr=thr, log=log)

    if img == 'image14.png':
        a[0:44, 0:130] = 255
        a[0:44, 540:818] = 255
        E.put(a, (36, 96, 73, 110), E.NEW_ACCOUNT, align='left', thr=thr)
        E.put(a, (36, 314, 88, 332), E.NEW_NAME, align='left', thr=thr)
        log.append('  manual: title whiteout + account + name')
    if img == 'image19.png':
        E.put(a, (190, 597, 378, 618), '2026-05-16 11:14:22', max_w=170, align='left', thr=thr)
        log.append('  manual: date field')
    if img == 'image46.png':
        # 顶部查询框里的值（白底深字）
        qb = L.bbox_of(L.dark_mask(a, 150), 100, 345, 15, 70)
        if qb:
            E.put(a, qb, E.NEW_ACCOUNT, max_w=210, align='left', thr=150)
            log.append('  manual: query input %s' % (qb,))
    if img == 'image22.png':
        # 弹窗遮罩下的灰色截图：整行重绘日期（两行），并改账号/姓名
        dates = [['2026-06-05', '13:26:38'], ['2026-06-05', '11:27:44'],
                 ['2026-05-16', '11:14:22'], ['2026-08-27', '11:03:32']]
        dband = L.cells_in_band(a, 770, 860, 0, a.shape[0] - 1, dark_thr=115)
        for box, dt in zip(dband, dates):
            L2.paint_lines(a, box, dt, thr=115)
        acct = [(890, 19, 962, 32), (890, 133, 962, 146), (890, 248, 962, 261), (890, 366, 962, 375)]
        name = [(972, 9, 1062, 33), (972, 123, 1062, 147), (972, 238, 1062, 262), (972, 352, 1062, 376)]
        for bx in acct:
            E.put(a, bx, E.NEW_ACCOUNT, max_w=95, thr=115)
        for bx in name:
            E.put(a, bx, E.NEW_NAME, max_w=90, thr=115)
        log.append('  manual(image22): dates=%d acct=%d name=%d' % (len(dband), len(acct), len(name)))

    if img == 'image13.png':
        L2.paint_text(a, (128, 37, 782, 75), E.BRAND, bold=True,
                      color=(0, 0, 0), bg=(255, 255, 255), max_w=600)
        log.append('  brand: title replaced')
    elif img in BRAND_NAVBAR:
        y0, y1 = (4, 62) if img == 'image19.png' else (4, 60)
        a[y0:y1, 1120:] = E.NAVY
        pil = Image.fromarray(a)
        d = ImageDraw.Draw(pil)
        d.text((18, (y0 + y1) // 2), E.BRAND, font=L.font(20, False),
               fill=(255, 255, 255), anchor='lm')
        d.text((a.shape[1] - 12, (y0 + y1) // 2), '用户：' + E.NEW_ACCOUNT,
               font=L.font(17, False), fill=(255, 255, 255), anchor='rm')
        a = np.asarray(pil)
        log.append('  brand: navbar + account')
    else:
        a = L2.brand_band(a)
        log.append('  brand: top band')

    Image.fromarray(a).save(os.path.join(OUT, img))
    return a


def main():
    which = sys.argv[1:] or ['image%d.png' % i for i in range(13, 47)]
    log = []
    for img in which:
        process(img, log)
    txt = '\n'.join(log)
    open(os.path.join(HERE, 'edit_log.txt'), 'w', encoding='utf-8').write(txt)
    print(txt)


if __name__ == '__main__':
    main()
