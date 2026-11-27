# -*- coding: utf-8 -*-
"""版面流式布局 v2：区域扣减 + 均衡网格 + 加入元素重排

设计要点
    1. 列数不再"能塞多少塞多少"：先按 max_cols 上限，再按行数回推做均衡
       （6 个元素 → 3×2；7 个 → 4+3；4 个 → 4×1）
    2. 行高统一：同一行元素等高；剩余高度按比例分给各行（不单独拉伸某张图）
    3. 溢出：先压图片（≥60%），再收间距，最后报告
"""
import sys

sys.path.insert(0, r'<LIB_ROOT>\tools')
from typeset import break_lines, line_height  # noqa: E402

PAGE_W, PAGE_H = 1280, 720
MARGIN = 40


def default_zones(nav=None, title_lines=1, footer=True):
    return {
        'page': (PAGE_W, PAGE_H),
        'margin': MARGIN,
        'nav': {'edge': nav, 'h': 64, 'w': 220} if nav else None,
        'title': {'lines': title_lines, 'pt': 32, 'gap': 16},
        'footer': {'h': 36} if footer else None,
    }


def content_area(z, page=None):
    pw, ph = page or z['page']
    m = z['margin']
    x, y, w, h = m, m, pw - 2 * m, ph - 2 * m
    nav = z.get('nav')
    if nav:
        if nav['edge'] == 'top':
            y += nav['h']; h -= nav['h']
        elif nav['edge'] == 'left':
            x += nav['w']; w -= nav['w']
    title = z.get('title')
    th = 0
    if title:
        th = round(line_height(title['pt']) * title['lines']) + title['gap']
        y += th; h -= th
    footer = z.get('footer')
    if footer:
        h -= footer['h']
    return {
        'x': x, 'y': y, 'w': max(w, 0), 'h': max(h, 0),
        'title_box': {'x': x, 'y': y - th, 'w': w, 'h': max(th - (title['gap'] if title else 0), 0)} if title else None,
        'footer_box': {'x': x, 'y': ph - m - footer['h'], 'w': w, 'h': footer['h']} if footer else None,
    }


def balanced_cols(n, max_cols=4):
    # 均衡列数：不出现最后一行只剩 1 个
    if n <= max_cols:
        return max(n, 1)
    rows = -(-n // max_cols)
    return -(-n // rows)


def solve(area, items, gap=(20, 16), max_cols=4, min_img_ratio=0.6):
    n = len(items)
    if n == 0:
        return {'cols': 0, 'rows': 0, 'rects': [], 'warnings': []}
    gx, gy = gap
    cols = balanced_cols(n, max_cols)
    rows = -(-n // cols)
    cw = (area['w'] - gx * (cols - 1)) / cols
    warnings = []

    # 每行基准高
    base_h, row_items = [], []
    for r in range(rows):
        chunk = items[r * cols:(r + 1) * cols]
        row_items.append(chunk)
        hs = []
        for it in chunk:
            if it['role'] == 'text':
                lines = break_lines(it.get('text', ''), cw - 24, it.get('pt', 20)) if it.get('text') else ['']
                hs.append(max(round(len(lines) * line_height(it.get('pt', 20))) + 24, it.get('min_h', 60)))
            else:
                hs.append(it.get('min_h', 120))
        base_h.append(max(hs) if hs else 0)

    # 行高统一：所有行取同一高度（多行时不出现高低不齐）
    if rows > 1:
        uniform = max(base_h)
        base_h = [uniform] * rows

    avail = area['h'] - gy * (rows - 1)
    total = sum(base_h)

    # 溢出：缩图片所在行
    if total > avail:
        over = total - avail
        img_rows = [i for i, chunk in enumerate(row_items) if any(it['role'] == 'image' for it in chunk)]
        for i in img_rows:
            if over <= 0:
                break
            floor_h = round(base_h[i] * min_img_ratio)
            cut = min(over, base_h[i] - floor_h)
            base_h[i] -= cut
            over -= cut
        if over > 0:
            warnings.append('内容超出 %.0fpx，建议减字或减图' % over)
        else:
            warnings.append('已压缩图片行高度以容纳内容')

    # 留白：按比例分配给所有行（元素整体铺满）
    total = sum(base_h)
    if total < avail:
        scale = avail / max(total, 1)
        base_h = [round(h * scale) for h in base_h]
        leftover = avail - sum(base_h)
        j = 0
        while leftover > 0 and base_h:
            base_h[j % len(base_h)] += 1
            leftover -= 1
            j += 1
        warnings.append('行高已铺满内容区')

    # 生成矩形
    rects, y = [], area['y']
    for r, chunk in enumerate(row_items):
        for c, it in enumerate(chunk):
            rects.append({'x': round(area['x'] + c * (cw + gx)), 'y': round(y),
                          'w': round(cw), 'h': int(base_h[r]), 'role': it['role']})
        y += base_h[r] + gy

    return {'cols': cols, 'rows': rows, 'rects': rects, 'warnings': warnings}


def add_text(items, text, pt=20, min_h=60):
    return items + [{'role': 'text', 'text': text, 'pt': pt, 'min_h': min_h}]


def add_image(items, min_h=120):
    return items + [{'role': 'image', 'min_h': min_h}]


if __name__ == '__main__':
    for nav in (None, 'top', 'left'):
        z = default_zones(nav=nav, title_lines=1, footer=True)
        a = content_area(z)
        print('nav=%-5s → 内容区 x=%d y=%d w=%d h=%d' % (str(nav), a['x'], a['y'], a['w'], a['h']))
    z = default_zones(nav='left', title_lines=1)
    a = content_area(z)
    for n in (2, 3, 4, 5, 6, 7, 8):
        r = solve(a, [{'role': 'text', 'text': '市井记忆未散', 'pt': 20}] * n)
        print('  %d 个元素 → %d 列 × %d 行' % (n, r['cols'], r['rows']))
