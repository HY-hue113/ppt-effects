# -*- coding: utf-8 -*-
"""成品后处理：补平滑（morph）+ 生成缺陷自检（整页复制导致的字母不位移等）。

用法：
  python postfix_deck.py --pptx out.pptx [--force] [--no-backup]

输出 JSON：
  {"ok": true, "pages": 96, "morphAdded": 87, "warnings": [{"code":"letters-static", ...}]}
"""
import argparse
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppt_morph import ensure_morph  # noqa: E402
from pptx import Presentation  # noqa: E402


def letter_signature(slide):
    letters = []
    for shape in slide.shapes:
        if not getattr(shape, 'has_text_frame', False):
            continue
        value = shape.text_frame.text.strip()
        if len(value) == 1 and value.isalpha() and (shape.top or 0) >= 0:
            letters.append(value + '@' + str(round((shape.left or 0) / 360000)))
    return '|'.join(sorted(letters))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--pptx', required=True)
    parser.add_argument('--force', action='store_true', help='已有非平滑切换的页也改成平滑')
    parser.add_argument('--no-backup', action='store_true')
    parser.add_argument('--expect-pages', type=int, default=0)
    parser.add_argument('--expect-letters', type=int, default=0)
    args = parser.parse_args()

    path = Path(args.pptx)
    if not path.is_file():
        print(json.dumps({'ok': False, 'error': 'file missing: ' + str(path)}, ensure_ascii=False))
        return 1

    prs = Presentation(str(path))
    added, skipped = ensure_morph(prs, force=args.force)

    signatures = []
    for index, slide in enumerate(prs.slides, start=1):
        signature = letter_signature(slide)
        if signature:
            signatures.append((index, signature))

    errors = []
    warnings = []

    # 平滑没补上（极少见）→ FAIL
    missing_morph = []
    for index, slide in enumerate(prs.slides, start=1):
        if slide._element.find('.//{http://schemas.microsoft.com/office/powerpoint/2015/09/main}morph') is None:
            missing_morph.append(index)
    if missing_morph:
        errors.append({'code': 'morph-missing', 'pages': missing_morph,
                       'message': '第 ' + ','.join(str(p) for p in missing_morph[:8]) + ' 页没有平滑切换'})

    # 图片占位没填 → FAIL
    placeholders = []
    for index, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not getattr(shape, 'has_text_frame', False):
                continue
            value = shape.text_frame.text.strip()
            if not value or not any(word in value for word in ('添加图片', '插入图片', '图片占位')):
                continue
            if shape._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip') is None:
                placeholders.append(index)
                break
    if placeholders:
        errors.append({'code': 'image-placeholder', 'pages': placeholders,
                       'message': '第 ' + ','.join(str(p) for p in placeholders[:8]) + ' 页还有「添加图片」占位没填'})
    # 元数据核对：页数 / 每页字母数（来自工作台的页序，不是猜）
    if args.expect_pages > 0 and len(prs.slides) != args.expect_pages:
        errors.append({'code': 'pages-mismatch', 'expected': args.expect_pages, 'actual': len(prs.slides),
                       'message': '页数不对：页序要求 %d 页，成品是 %d 页' % (args.expect_pages, len(prs.slides))})
    if args.expect_letters > 0:
        bad_pages = []
        for index, slide in enumerate(prs.slides, start=1):
            tops = []
            for shape in slide.shapes:
                if not getattr(shape, 'has_text_frame', False):
                    continue
                value = shape.text_frame.text.strip()
                if len(value) == 1 and value.isalpha():
                    tops.append(shape.top or 0)
            if not tops:
                continue
            # 有字母被移出画布（散开/封面）→ 这是效果本身的布局，跳过不核对
            if any(top < 0 for top in tops):
                continue
            if len(tops) != args.expect_letters:
                bad_pages.append((index, len(tops)))
        if bad_pages:
            errors.append({'code': 'letters-count', 'expected': args.expect_letters,
                           'pages': [page for page, _count in bad_pages],
                           'message': '字母数不对：应每页 %d 个，实际 ' % args.expect_letters
                                      + '、'.join('第%d页%d个' % pair for pair in bad_pages[:6])})
    if len(signatures) >= 3:
        counts = {}
        for _index, signature in signatures:
            counts[signature] = counts.get(signature, 0) + 1
        most = max(counts.values())
        if most >= 6 and len(counts) == 1:
            warnings.append({
                'code': 'letters-static',
                'pages': most,
                'message': '这 %d 页的字母对象与位置完全一致，平滑切换不会产生字形位移；像是整页复制，应该用效果构建脚本按图数重排' % most,
            })

    if not args.no_backup:
        backup = path.with_suffix(path.suffix + '.bak')
        if not backup.exists():
            shutil.copy2(path, backup)

    temp = path.with_name(path.stem + '-postfix.pptx')
    prs.save(str(temp))
    temp.replace(path)

    print(json.dumps({
        'ok': not errors,
        'pages': len(prs.slides),
        'morphAdded': added,
        'morphSkipped': skipped,
        'errors': errors,
        'warnings': warnings,
    }, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
