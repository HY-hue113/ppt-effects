#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把「整块在画布外」的根级 <g> 包成 morph staging 组，让质量门放行。

背景：ppt-master 的 svg_quality_checker 会把越出 viewBox 的 <text>/<image> 判 ERROR。
但「整块停在画布外等命」（效果 5 上滚、效果 20 照片待命等）是合法设计，官方豁免是：

    <g id="stage-bottom" data-pptx-morph-staging="true" data-pptx-bounds="x y w h">

三条硬约束（不满足照样 ERROR）：
  1. 该组必须**完全**在画布外（与 viewBox 不相交）
  2. 组内内容的估算范围必须落在 data-pptx-bounds 内（容差 5%）
  3. 必须有稳定的 id，且不能与 data-pptx-layer / data-pptx-placeholder 同时出现

用法：
    python tools/stage_offscreen.py <svg目录或svg文件> [--check]

脚本只处理 data-pptx-frame 已经完全在画布外的根级 <g>；已经是 staging 的会跳过，
重复运行结果不变（幂等）。处理完建议再跑一次质量门确认。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lxml import etree

NS = 'http://www.w3.org/2000/svg'
STAGING = 'data-pptx-morph-staging'
FRAME = 'data-pptx-frame'
BOUNDS = 'data-pptx-bounds'


def parse_frame(g):
    raw = g.get(FRAME)
    if not raw:
        return None
    try:
        x, y, w, h = (float(v) for v in raw.split())
    except Exception:
        return None
    return x, y, w, h


def side_of(rect, W, H):
    """完全在画布外时返回所在边，否则 None。"""
    x, y, w, h = rect
    if y + h <= 0:
        return 'top'
    if y >= H:
        return 'bottom'
    if x + w <= 0:
        return 'left'
    if x >= W:
        return 'right'
    return None


def canvas_of(root):
    vb = root.get('viewBox')
    if vb:
        try:
            _, _, w, h = (float(v) for v in vb.replace(',', ' ').split())
            return w, h
        except Exception:
            pass
    try:
        return float(root.get('width')), float(root.get('height'))
    except Exception:
        return 1280.0, 720.0


def bounds_for(rects, side, W, H, pad=600.0):
    """向外扩，保证与画布完全不相交。"""
    x0 = min(r[0] for r in rects) - pad
    x1 = max(r[0] + r[2] for r in rects) + pad
    y0 = min(r[1] for r in rects) - pad
    y1 = max(r[1] + r[3] for r in rects) + pad
    if side == 'top':
        y1 = min(y1, -8.0)
    elif side == 'bottom':
        y0 = max(y0, H + 8.0)
    elif side == 'left':
        x1 = min(x1, -8.0)
    elif side == 'right':
        x0 = max(x0, W + 8.0)
    return x0, y0, x1 - x0, y1 - y0


def process(path: Path, check_only=False):
    tree = etree.parse(str(path))
    root = tree.getroot()
    W, H = canvas_of(root)

    groups = [g for g in root if etree.QName(g).localname == 'g']
    buckets = {}
    for g in groups:
        if g.get(STAGING) == 'true':
            continue
        rect = parse_frame(g)
        if not rect:
            continue
        side = side_of(rect, W, H)
        if side:
            buckets.setdefault(side, []).append((g, rect))

    if check_only:
        return {k: len(v) for k, v in buckets.items()}, None

    made = []
    for side, items in buckets.items():
        gid = 'stage-%s' % side
        if root.find(f'{{{NS}}}g[@{STAGING}="true"]') is not None:
            used = {g.get('id') for g in root if etree.QName(g).localname == 'g'}
            n = 2
            base = gid
            while gid in used:
                gid = '%s-%d' % (base, n)
                n += 1
        bx, by, bw, bh = bounds_for([r for _, r in items], side, W, H)
        wrapper = etree.SubElement(root, f'{{{NS}}}g')
        wrapper.set('id', gid)
        wrapper.set(STAGING, 'true')
        wrapper.set(BOUNDS, '%.1f %.1f %.1f %.1f' % (bx, by, bw, bh))
        for g, _ in items:
            root.remove(g)
            wrapper.append(g)
        made.append((gid, len(items), (bx, by, bw, bh)))

    if made:
        tree.write(str(path), encoding='utf-8', xml_declaration=True)
    return {}, made


def main() -> int:
    ap = argparse.ArgumentParser(description='给整块出画的根级 <g> 加 morph staging 豁免')
    ap.add_argument('target', help='SVG 目录或单个 .svg')
    ap.add_argument('--check', action='store_true', help='只报告，不修改')
    args = ap.parse_args()

    tgt = Path(args.target)
    files = sorted(tgt.glob('*.svg')) if tgt.is_dir() else [tgt]
    if not files:
        print('没有找到 SVG:', tgt)
        return 1

    total = 0
    for f in files:
        pending, made = process(f, args.check)
        if args.check:
            if pending:
                print('%-28s 待处理: %s' % (f.name, ', '.join('%s=%d' % kv for kv in pending.items())))
                total += sum(pending.values())
        else:
            if made:
                for gid, n, b in made:
                    print('%-28s %-14s 收 %2d 个   bounds=%.0f %.0f %.0f %.0f' % (f.name, gid, n, *b))
                total += sum(n for _, n, _ in made)
            else:
                print('%-28s 无需处理' % f.name)

    print('\n合计 %d 个出画组%s' % (total, '（未修改）' if args.check else '已加豁免'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
