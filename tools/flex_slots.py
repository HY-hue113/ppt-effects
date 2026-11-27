# -*- coding: utf-8 -*-
"""柔性分配：同一页的内容槽位可按文字量在「文框 / 图形框」之间切换

输入：
    slots   页面槽位（文框/图形框/标题/导航/页脚）
    text_chars  用户准备放的字数
    images      准备放的图片数
输出：
    plan = [{'role': '文框'|'图形框'|'标题'|'导航'|'页脚', 'x','y','w','h',
             'points': [...], 'curves': [...]}]
    points 为实际顶点，curves 为每条边的二次曲线控制点（null=直边）；
    改槽位包围盒时同步缩放 points/curves。

分配规则：
    先按 images 预留图片槽位 → 文字按容量占用最少槽位 → 剩余给图片
    图片少于剩余槽位时合并相邻槽（图更大）

元素自适应边界（2026-11-19 用户确认）：
    标题     可变：宽 380–1160、高 56→96（1→2 行）      不可变：顶部对齐
    横向导航 可变：高 64–72（厚度）                      不可变：宽（长短）
    竖向导航 可变：宽 210–230（厚度）                    不可变：高（长短）
    内容槽   可变：角色互换、合并、放大、下移避让        不可变：不出页面
"""

import sys

sys.path.insert(0, r'<LIB_ROOT>\tools')
from typeset import break_lines  # noqa: E402

TEXT_CAP = 24          # 每格每行容纳字数（约 20pt，宽 300px）
LINE_H = 36


def chars_capacity(slot):
    """该槽位放 20pt 文字能装多少字"""
    cols = max(1, int((slot['w'] - 24) / 22))
    rows = max(1, int((slot['h'] - 24) / LINE_H))
    return cols * rows


def shape_bbox(slot):
    points = slot.get('points') or []
    if points:
        xs = [p['x'] for p in points]
        ys = [p['y'] for p in points]
        return min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)
    return slot['x'], slot['y'], slot['w'], slot['h']


def set_bbox(slot, x, y, w, h):
    """改包围盒时同步缩放 points/curves，避免旧四边形与新 x/y/w/h 脱节。"""
    old_x, old_y, old_w, old_h = shape_bbox(slot)
    points = slot.get('points')
    if points:
        sx = w / max(1.0, old_w)
        sy = h / max(1.0, old_h)
        slot['points'] = [
            {'x': x + (p['x'] - old_x) * sx, 'y': y + (p['y'] - old_y) * sy}
            for p in points
        ]
        curves = slot.get('curves')
        if curves:
            slot['curves'] = [
                None if c is None else {
                    'x': x + (c['x'] - old_x) * sx,
                    'y': y + (c['y'] - old_y) * sy,
                }
                for c in curves
            ]
    slot.update({'x': x, 'y': y, 'w': w, 'h': h})


def set_rect(slot, x, y, w, h):
    """合并槽位时重建成矩形，避免把多个不同形状硬拉成一个。"""
    slot['points'] = [
        {'x': x, 'y': y},
        {'x': x + w, 'y': y},
        {'x': x + w, 'y': y + h},
        {'x': x, 'y': y + h},
    ]
    slot['curves'] = [None, None, None, None]
    slot.update({'x': x, 'y': y, 'w': w, 'h': h})


def plan_page(slots, text_chars, images=0):
    content = [s for s in slots if s['role'] in ('文框', '图形框')]
    others = [s for s in slots if s['role'] not in ('文框', '图形框')]
    if not content:
        return list(slots)

    # 按阅读顺序（先上后下、先左后右）
    ordered = sorted(content, key=lambda s: (s['y'], s['x']))
    caps = [chars_capacity(s) for s in ordered]

    # 0) 先给图片预留槽位（至少留 images 个，但不抢光文字位）
    img_reserve = min(max(images, 0), max(len(ordered) - 1, 0))
    max_text_slots = max(1, len(ordered) - img_reserve)

    # 1) 用最少的前几个槽位装下文字
    text_n = 0
    acc = 0
    while text_n < max_text_slots and acc < text_chars:
        acc += caps[text_n]
        text_n += 1
    text_n = max(text_n, 0)
    if text_chars <= 0:
        text_n = 0
    # 至少要给图片留一格（除非没有图也没有剩余槽）
    if text_n >= len(ordered) and (images > 0 or len(ordered) > 1):
        text_n = len(ordered) - 1

    text_slots = ordered[:text_n]
    rest = ordered[text_n:]

    # 2) 剩余槽位给图片：**一律保留版式原生图片框，不合并**
    #    历史做法是把多余的图片槽并成一个并集矩形，结果会跨行/跨角色生成
    #    1198×552 这种巨框（p18 五格 → 2 格），既不是版式原本的框，也违反
    #    “未选图片框时也要填版式自带图形框”。images 只表示用户准备了几张图，
    #    不再改动槽位几何；图不够时由生成端循环使用或先问用户。
    img_slots = [dict(s) for s in rest]

    plan = [
        {**s, 'role': s.get('role', '文框') if s.get('points') else '文框'}
        for s in text_slots
    ] + [
        {**s, 'role': s.get('role', '图形框') if s.get('points') else '图形框'}
        for s in img_slots
    ]
    plan.extend(others)
    plan.sort(key=lambda s: (s['y'], s['x']))
    return plan


def adapt_title(plan, title_text, max_w=1160, min_w=380):
    """只让主标题按文字量自适应；卡片里的多个小标题保持原框。"""
    titles = [s for s in plan if s['role'] == '标题']
    primary = min(titles, key=lambda s: (s['y'], s['x'])) if titles else None
    if primary and primary['y'] > 160:
        primary = None
    for s in ([primary] if primary else []):
        w = min(max_w, max(min_w, 22 * len(title_text) + 80), max(min_w, 1280 - s['x'] - 20))
        h = 56 if len(title_text) <= 26 else 96
        set_bbox(s, s['x'], s['y'], w, h)
        # 导航：只调厚度，不调长短
        #   横向导航（宽>高×3）→ 宽度锁死，只允许改高度（48–80）
        #   竖向导航（高>宽×2）→ 高度锁死，只允许改宽度（160–260）
        for n in plan:
            if n['role'] != '导航':
                continue
            if n['w'] >= n['h'] * 3:
                set_bbox(n, n['x'], n['y'], n['w'], 64 if s['h'] <= 56 else 72)
            elif n['h'] >= n['w'] * 2:
                set_bbox(n, n['x'], n['y'], 210 if s['h'] <= 56 else 230, n['h'])

    # 标题与内容重叠 → 内容整体下移（并压缩高度，保证不出页）
    title = primary
    if title:
        delta = 0
        for s2 in plan:
            if s2['role'] in ('文框', '图形框') and s2['y'] < title['y'] + title['h']:
                delta = max(delta, title['y'] + title['h'] + 12 - s2['y'])
        if delta:
            for s2 in plan:
                if s2['role'] in ('文框', '图形框'):
                    new_y = s2['y'] + delta
                    new_h = s2['h']
                    bottom = 720 - 56
                    if new_y + new_h > bottom:
                        new_h = max(80, bottom - new_y)
                    if new_y > bottom - 80:
                        new_y = bottom - 80
                        new_h = 80
                    set_bbox(s2, s2['x'], new_y, s2['w'], new_h)
    return plan


def scenario(slots, name, text_chars, images, title='项目标题'):
    p = adapt_title(plan_page(slots, text_chars, images), title)
    texts = [s for s in p if s['role'] == '文框']
    imgs = [s for s in p if s['role'] == '图形框']
    print('%-10s 字数 %-4d 图 %d → 文框 %d 个，图形框 %d 个，标题宽 %s'
          % (name, text_chars, images, len(texts), len(imgs),
             next((s['w'] for s in p if s['role'] == '标题'), '-')))
    return p


# ---------------- 用户显式指定优先 ----------------
def apply_overrides(plan, overrides):
    """按用户指定改写分配结果（用户需求优先于自动分配）

    overrides 支持两种写法：
        1) 按槽位来源：{'from': 'row1', 'role': '图形框'}
        2) 按序号（0 基，按当前 plan 顺序）：{'index': 0, 'role': '文框'}
        3) 按区域：{'x': 60, 'y': 120, 'w': 400, 'h': 200, 'role': '图形框'}
    未匹配到的 override 会被记录到 plan['_unmatched']（供上层提示用户）。
    """
    unmatched = []
    for o in overrides or []:
        hit = None
        if 'from' in o:
            hit = next((s for s in plan if s.get('from') == o['from']), None)
        elif 'index' in o and 0 <= o['index'] < len(plan):
            hit = plan[o['index']]
        elif all(k in o for k in ('x', 'y', 'w', 'h')):
            hit = next((s for s in plan
                        if abs(s['x'] - o['x']) <= 2 and abs(s['y'] - o['y']) <= 2
                        and abs(s['w'] - o['w']) <= 2 and abs(s['h'] - o['h']) <= 2), None)
        if hit is None:
            unmatched.append(o)
            continue
        hit['role'] = o['role']
        if o.get('w') or o.get('h'):
            set_bbox(hit, hit['x'], hit['y'], o.get('w') or hit['w'], o.get('h') or hit['h'])
    return {'plan': plan, 'unmatched': unmatched}


def scenario_with_overrides(slots, name, text_chars, images, overrides, title='项目标题'):
    """自动分配 + 用户指定覆盖"""
    plan = scenario(slots, name, text_chars, images, title)
    plan = [s for s in plan if '_unmatched' not in s]
    return apply_overrides(plan, overrides)
