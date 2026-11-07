#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""套用效果前校验素材：数量、画幅、方向、明暗/白底前提。

素材要求写在每个模块的 spec.json -> materials 里（有 variants 时按 --variant 取）。
本脚本只做判定与告知，不改任何文件；判定口径也写在 spec 里，不是硬编码。

用法:
    python tools/check_materials.py 效果1 照片.jpg
    python tools/check_materials.py 效果9 ./素材文件夹/
    python tools/check_materials.py 效果2 --variant long *.jpg
    python tools/check_materials.py --list
    python tools/check_materials.py 效果7 a.jpg --json

退出码:
    0  可以继续（可能有 [待确认] 项，仍要先跟用户说）
    1  阻断（缺件、打不开、比例严重不符）
    2  用法错误（找不到效果、没给素材）
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:  # pragma: no cover
    raise SystemExit('需要 Pillow：pip install Pillow')

ROOT = Path(__file__).resolve().parent.parent
EFFECTS_DIR = ROOT / 'effects'
IMAGE_SUFFIXES = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff', '.gif'}

OK, WARN, BLOCK = 'ok', 'warn', 'block'
MARK = {OK: '[OK]', WARN: '[!] ', BLOCK: '[X] '}

# 比例偏差门槛：容差内算合格；超过容差算"要裁/要拉"；超过 25% 视为明显变形
SEVERE_DEVIATION = 0.25


def die(msg: str, code: int = 2) -> None:
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def load_spec(effect_dir: Path) -> dict:
    return json.loads((effect_dir / 'spec.json').read_text('utf-8'))


def iter_effects() -> list[Path]:
    return sorted(d for d in EFFECTS_DIR.iterdir() if d.is_dir() and d.name[:1].isdigit())


def resolve_effect(arg: str) -> Path:
    p = Path(arg)
    if p.is_dir() and (p / 'spec.json').exists():
        return p
    key = arg.strip().lower()
    for d in iter_effects():
        spec = load_spec(d)
        names = {spec.get('id', '').lower(), spec.get('id', '').replace('效果', '').lower(),
                 d.name.lower(), d.name[:3].lower(), Path(spec.get('slug', '')).name.lower()}
        if key in names:
            return d
    for d in iter_effects():          # 模糊：包含关系
        if key and (key in d.name.lower() or key in load_spec(d).get('id', '').lower()):
            return d
    die('找不到效果：%s（用 --list 看全部）' % arg)


def materials_of(spec: dict, variant: str | None) -> tuple[dict, str]:
    mats = spec.get('materials') or {}
    variants = mats.get('variants')
    if variants:
        if not variant:
            die('%s 有多个档位（%s），请用 --variant 指定' % (spec.get('id'), ' / '.join(variants)))
        if variant not in variants:
            die('未知档位：%s（可选 %s）' % (variant, ' / '.join(variants)))
        return variants[variant], variant
    return mats, ''


def collect_images(args: list[str]) -> list[Path]:
    out: list[Path] = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            found = sorted(f for f in p.iterdir() if f.is_file() and f.suffix.lower() in IMAGE_SUFFIXES)
            if not found:
                die('文件夹里没有图片：%s' % a)
            out.extend(found)
        elif p.is_file():
            out.append(p)
        else:
            die('文件不存在：%s' % a)
    return out


def probe(path: Path) -> dict:
    with Image.open(path) as im:
        w, h = im.size
        has_alpha_channel = im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info)
        alpha_ratio = 0.0
        if has_alpha_channel:
            a = im.convert('RGBA').getchannel('A').resize((64, 64))
            ap = a.load()
            alpha_ratio = sum(1 for x in range(64) for y in range(64) if ap[x, y] < 16) / 4096
        small = im.convert('RGB').resize((64, 64))
        px = small.load()
        total = 0.0
        for x in range(64):
            for y in range(64):
                r, g, b = px[x, y][:3]
                total += 0.299 * r + 0.587 * g + 0.114 * b
        luma = total / 4096
        border = [px[x, 0] for x in range(64)] + [px[x, 63] for x in range(64)] \
               + [px[0, y] for y in range(64)] + [px[63, y] for y in range(64)]
        white = sum(1 for r, g, b in (p[:3] for p in border) if min(r, g, b) >= 230) / len(border)
    aspect = w / h
    orientation = 'landscape' if aspect > 1.05 else ('portrait' if aspect < 0.95 else 'square')
    return {'path': str(path), 'name': path.name, 'w': w, 'h': h,
            'aspect': round(aspect, 3), 'orientation': orientation,
            'luma': round(luma, 1), 'white_ratio': round(white, 2),
            'has_alpha': has_alpha_channel, 'alpha_ratio': round(alpha_ratio, 2)}


def slot_orientations(aspects: list[float]) -> set[str]:
    out = set()
    for a in aspects:
        out.add('landscape' if a > 1.05 else ('portrait' if a < 0.95 else 'square'))
    return out


def check_images(images: list[dict], mats: dict, issues: list[dict]) -> None:
    aspects = [float(a) for a in mats.get('aspects') or []]
    tol = float(mats.get('aspect_tolerance', 0.06))
    want_orient = mats.get('orientation', 'any')
    allowed = slot_orientations(aspects) if aspects else set()

    for img in images:
        verdict, note = OK, ''
        # 画幅
        if aspects:
            nearest = min(aspects, key=lambda s: abs(img['aspect'] - s))
            dev = abs(img['aspect'] - nearest) / nearest
            img['nearest_slot'] = nearest
            img['deviation'] = round(dev, 3)
            if dev > SEVERE_DEVIATION:
                if mats.get('crop_allowed'):
                    cut = (1 - min(img['aspect'], nearest) / max(img['aspect'], nearest)) * 100
                    verdict, note = WARN, '画幅 %.2f 对槽 %.2f：会居中裁掉约 %.0f%%，不拉伸' % (
                        img['aspect'], nearest, cut)
                    issues.append({'level': WARN, 'code': 'aspect', 'message': '%s %s' % (img['name'], note)})
                else:
                    verdict, note = BLOCK, '画幅 %.2f，最近槽 %.2f，偏 %.0f%%：会被明显拉伸' % (img['aspect'], nearest, dev * 100)
                    issues.append({'level': BLOCK, 'code': 'aspect', 'message': '%s %s' % (img['name'], note)})
            elif dev > tol:
                verdict, note = WARN, '画幅 %.2f，最近槽 %.2f（偏 %.0f%%）：要裁切或轻拉伸，先问用户' % (img['aspect'], nearest, dev * 100)
                issues.append({'level': WARN, 'code': 'aspect', 'message': '%s %s' % (img['name'], note)})
            else:
                note = '画幅 %.2f 对槽 %.2f' % (img['aspect'], nearest)
        # 方向
        if allowed and img['orientation'] not in allowed and want_orient != 'any':
            if verdict == OK:
                verdict = WARN
            note = (note + '；' if note else '') + '方向 %s，槽位只要 %s' % (img['orientation'], '/'.join(sorted(allowed)))
            issues.append({'level': WARN, 'code': 'orientation',
                           'message': '%s 方向 %s，槽位只要 %s' % (img['name'], img['orientation'], '/'.join(sorted(allowed)))})
        # 明暗
        cap = (mats.get('checks') or {}).get('max_mean_luma')
        if cap is not None and img['luma'] > cap:
            if verdict == OK:
                verdict = WARN
            note = (note + '；' if note else '') + '平均亮度 %.0f 超过 %.0f，白字会看不清' % (img['luma'], cap)
            issues.append({'level': WARN, 'code': 'luma',
                           'message': '%s 偏亮（%.0f > %.0f）' % (img['name'], img['luma'], cap)})
        # 白底
        floor = (mats.get('checks') or {}).get('min_white_border')
        prepared = bool((mats.get('checks') or {}).get('accept_alpha'))
        if floor is not None and img['white_ratio'] < floor:
            if prepared and img.get('alpha_ratio', 0) >= 0.08:
                # 已经是抠好的透明图：这就是"处理过的成品输入"，直接可用
                note = (note + '；' if note else '') + '透明背景 %.0f%%，已是处理好的主体图' % (img['alpha_ratio'] * 100)
            else:
                hard = bool((mats.get('checks') or {}).get('white_bg_blocking'))
                if verdict != BLOCK:
                    verdict = BLOCK if hard else WARN
                note = (note + '；' if note else '') + '边缘白底占比 %.2f 低于 %.2f，%s' % (
                    img['white_ratio'], floor, '抽不出线稿' if hard else '抠图会带脏边')
                issues.append({'level': BLOCK if hard else WARN, 'code': 'white_bg',
                               'message': '%s 不是干净白底（%.2f < %.2f）%s' % (
                                   img['name'], img['white_ratio'], floor,
                                   '：这个效果靠亮度反推 alpha，非白底直接做不出来；已抠好的透明 PNG 可以直接给' if hard else '')})

        # 文字对比：底图太暗/太亮时，标题会看不见（判定阈值写在 spec 里）
        tc = (mats.get('checks') or {}).get('text_contrast') or {}
        need = tc.get('min_mean_luma')
        ceiling = tc.get('max_mean_luma')
        if need is not None and img['luma'] < need:
            if verdict == OK:
                verdict = WARN
            is_black = str(tc.get('title_color', '000000')).lower().startswith('0')
            fix = tc.get('fix', '换亮图，或把标题改色 / 加遮罩')
            msg = ('%s 偏暗（亮度 %.0f < %s）：%s 的标题是%s色，压上去会看不见 —— %s'
                   % (img['name'], img['luma'], need, tc.get('where', '本效果'),
                      '黑' if is_black else '白', fix))
            note = (note + '；' if note else '') + '亮度 %.0f 低于 %s，%s色标题会看不清' % (
                img['luma'], need, '黑' if is_black else '白')
            issues.append({'level': WARN, 'code': 'text_contrast', 'message': msg})
        if ceiling is not None and img['luma'] > ceiling:
            if verdict == OK:
                verdict = WARN
            note = (note + '；' if note else '') + '亮度 %.0f 高于 %s，浅色标题会看不清' % (img['luma'], ceiling)
            issues.append({'level': WARN, 'code': 'text_contrast',
                           'message': '%s 偏亮（亮度 %.0f > %s）：白字会看不清' % (img['name'], img['luma'], ceiling)})

        img['verdict'] = verdict
        img['note'] = note


def recommended_range(mats: dict) -> tuple[int | None, int | None]:
    """materials.recommended：4 / [4, 5] / "4-5" 三种写法都收。"""
    rec = mats.get('recommended')
    if rec is None:
        return None, None
    if isinstance(rec, str):
        m = re.match(r'\s*(\d+)\s*(?:[-–~]\s*(\d+))?\s*$', rec)
        if not m:
            return None, None
        lo = int(m.group(1)); return lo, int(m.group(2)) if m.group(2) else lo
    if isinstance(rec, (list, tuple)) and rec:
        lo = int(rec[0]); return lo, int(rec[1]) if len(rec) > 1 else lo
    if isinstance(rec, int):
        return rec, rec
    return None, None


def fmt_range(lo: int, hi: int | None) -> str:
    if hi is None or hi == lo:
        return '%d' % lo
    return '%d–%d' % (lo, hi)


def check_count(n: int, mats: dict, issues: list[dict]) -> str:
    lo, hi = int(mats.get('min', 0)), int(mats.get('max', 0))
    policy = mats.get('deficit', 'none')
    if n < lo:
        if policy == 'none':
            msg = '要 %s，只给了 %d 张 —— 这个效果缺图做不出来' % (lo if lo == hi else '≥%d' % lo, n)
            issues.append({'level': BLOCK, 'code': 'short', 'message': msg})
        else:
            gap = lo - n
            tail = {'keep-old': '，缺的 %d 个槽位会沿用参考件里的照片（成品会混入你没给的图）' % gap,
                    'fewer-pages': '，会按 %d 张出页（结构自动缩小，不留旧图）' % n,
                    'repeat': '，缺的 %d 个槽位会重复使用你给的这 %d 张（不会混入旧图）' % (gap, n)}[policy]
            issues.append({'level': WARN, 'code': 'short', 'message': '要 ≥%d 张，只给了 %d 张%s' % (lo, n, tail)})
    if hi and n > hi:
        issues.append({'level': WARN, 'code': 'over',
                       'message': '最多用 %d 张，多出的 %d 张会闲置' % (hi, n - hi)})
    rec_lo, rec_hi = recommended_range(mats)
    if rec_lo and lo <= n < rec_lo:
        note = mats.get('recommended_note') or ''
        issues.append({'level': WARN, 'code': 'below-recommended',
                       'message': '推荐 %s 张，只给了 %d 张%s' % (fmt_range(rec_lo, rec_hi), n,
                                                             '：' + note if note else '')})
    return policy


def render(effect_dir: Path, spec: dict, variant: str, mats: dict, images: list[dict],
           issues: list[dict]) -> dict:
    text = mats.get('requirement', '')
    rng = '%d' % mats['min'] if mats.get('min') == mats.get('max') else '%s–%s' % (mats.get('min'), mats.get('max'))
    if mats.get('max') == 0:
        rng = '0'
    policy_hint = {'none': '缺了做不出来', 'keep-old': '缺了会留旧图',
                   'fewer-pages': '少图会自动少页', 'repeat': '缺了会重复用图'}.get(mats.get('deficit', 'none'), '')
    print('=' * 68)
    print('%s %s%s' % (spec.get('id'), spec.get('name'), '（%s 版）' % variant if variant else ''))
    print('数量要求：%s 张%s（%s）' % (rng, '，推荐 %s' % fmt_range(*recommended_range(mats)) if recommended_range(mats)[0] else '', policy_hint))
    if text:
        print('素材要求：%s' % text)
    if mats.get('verified') is False:
        print('           （槽位数按 recipe/index 推断，未逐槽核对）')
    if mats.get('note'):
        print('注：%s' % mats['note'])
    if mats.get('aspects'):
        uniq = sorted({round(float(a), 2) for a in mats['aspects']})
        print('槽位画幅：%s（容差 ±%.0f%%）' % (' / '.join('%.2f' % a for a in uniq),
                                        float(mats.get('aspect_tolerance', 0.06)) * 100))
    print('实际给定：%d 张' % len(images))
    print('-' * 68)
    for img in images:
        print('%s %-24s %5d×%-5d %s %-9s 亮度%-5.1f %s' % (
            MARK[img['verdict']], img['name'][:24], img['w'], img['h'],
            '%.2f' % img['aspect'], img['orientation'], img['luma'], img['note']))
    if not images:
        print('（没有可用的图片）')
    print('-' * 68)

    blocking = [i for i in issues if i['level'] == BLOCK]
    warning = [i for i in issues if i['level'] == WARN]
    for i in blocking:
        print('[X]  %s' % i['message'])
    for i in warning:
        print('[!]  %s' % i['message'])
    verdict = 'BLOCK' if blocking else ('WARN' if warning else 'PASS')
    print()
    if verdict == 'BLOCK':
        print('结论：BLOCK —— 先补素材或跟用户确认，不要硬套。')
    elif verdict == 'WARN':
        print('结论：可以先做，但这几条必须先跟用户说，不能默默继续。')
    else:
        print('结论：PASS —— 数量与画幅都在槽位内，可以开始套用。')
    return {'verdict': verdict, 'issues': issues, 'images': images}


def cmd_list() -> None:
    def rng(m: dict) -> str:
        lo, hi = m.get('min', 0), m.get('max', 0)
        base = str(lo) if lo == hi else '%s–%s' % (lo, hi)
        rec = recommended_range(m)[0]
        return '%s（荐%s）' % (base, fmt_range(*recommended_range(m))) if rec else base

    rows = []
    for d in iter_effects():
        spec = load_spec(d)
        mats = spec.get('materials') or {}
        variants = mats.get('variants')
        if variants:
            for k, v in variants.items():
                rows.append((spec.get('id', ''), '%s（%s）' % (spec.get('name', ''), k),
                             rng(v), v.get('requirement', '')))
        else:
            rows.append((spec.get('id', ''), spec.get('name', ''), rng(mats) if mats else '-',
                         mats.get('requirement', '（未声明）')))
    idw = max(len(r[0]) for r in rows)
    namew = max(len(r[1]) for r in rows)
    for eid, name, cnt, req in rows:
        print('%-*s  %-*s  %-5s  %s' % (idw, eid, namew, name, cnt, req))


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8')
        except Exception:
            pass

    ap = argparse.ArgumentParser(description='套用效果前校验素材（数量/画幅/方向/明暗），不改文件')
    ap.add_argument('effect', nargs='?', help='效果代号（效果1 / 1 / 009_金边弧线）或模块目录')
    ap.add_argument('images', nargs='*', help='图片文件或文件夹')
    ap.add_argument('--variant', default=None, help='多档效果的档位，如 long')
    ap.add_argument('--json', action='store_true', help='输出 JSON（给脚本/自动化读）')
    ap.add_argument('--list', action='store_true', help='列出所有效果的素材要求')
    args = ap.parse_args()

    if args.list:
        cmd_list()
        return 0
    if not args.effect:
        die('要给效果代号；用 --list 查看全部')
    if not args.images:
        die('要给素材文件或文件夹')

    effect_dir = resolve_effect(args.effect)
    spec = load_spec(effect_dir)
    mats, variant = materials_of(spec, args.variant)
    if not mats:
        die('%s 的 spec.json 里没有 materials 段' % spec.get('id'))
    if mats.get('max') == 0:
        print('%s 不依赖照片（纯矢量），素材校验跳过。' % spec.get('id'))
        return 0

    paths = collect_images(args.images)
    issues: list[dict] = []
    images = [probe(p) for p in paths]
    check_count(len(images), mats, issues)
    check_images(images, mats, issues)

    if args.json:
        import contextlib
        import io
        payload = {'effect': spec.get('id'), 'name': spec.get('name'), 'variant': variant,
                   'requirement': mats.get('requirement', ''),
                   'required': {'min': mats.get('min'), 'max': mats.get('max')},
                   'recommended': list(recommended_range(mats)) if recommended_range(mats)[0] else None,
                   'given': len(images), 'deficit_policy': mats.get('deficit', 'none')}
        with contextlib.redirect_stdout(io.StringIO()):
            payload.update(render(effect_dir, spec, variant, mats, images, issues))
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        render(effect_dir, spec, variant, mats, images, issues)

    return 1 if any(i['level'] == BLOCK for i in issues) else 0


if __name__ == '__main__':
    raise SystemExit(main())
