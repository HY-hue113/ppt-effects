# -*- coding: utf-8 -*-
"""整册换图：按文档顺序把图片槽（背景图 / p:pic / 图片填充形状，含组合内）依次换成给定图片。

用法：
  python refill_images.py --pptx in.pptx --images a.jpg,b.jpg [--page 3] --out out.pptx

说明：
  - 支持组合（p:grpSp）里的 p:pic 与 a:blipFill 形状（效果22 的照片都在组合里）；
  - 只改图片关系 blob，不动几何、动画、切换；
  - 图片数不足时循环使用。
"""
import argparse
import shutil
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

A = qn('a:blip')
EMBED = qn('r:embed')


def slot_elements(slide):
    """返回本页所有可换图的 XML 元素（按文档顺序，含组合内部）。"""
    root = slide._element
    slots = []

    bg = root.find('.//' + qn('p:bg'))
    if bg is not None:
        blip = bg.find('.//' + A)
        if blip is not None:
            slots.append(('bg', bg, blip))

    for pic in root.iter(qn('p:pic')):
        blip = pic.find('.//' + A)
        if blip is not None:
            slots.append(('pic', pic, blip))

    for shape in root.iter(qn('p:sp')):
        fill = shape.find('.//' + qn('a:blipFill'))
        if fill is None:
            continue
        blip = fill.find('.//' + A)
        if blip is not None:
            slots.append(('fill', shape, blip))
    return slots


def frame_of(container):
    xfrm = container.find('.//' + qn('a:xfrm'))
    if xfrm is None:
        return None
    ext = xfrm.find(qn('a:ext'))
    if ext is None:
        return None
    try:
        cx = int(ext.get('cx'))
        cy = int(ext.get('cy'))
        return cx, cy
    except (TypeError, ValueError):
        return None


def cover_src_rect(container, image_path: Path):
    frame = frame_of(container)
    if not frame:
        return None
    cx, cy = frame
    if not cx or not cy:
        return None
    frame_ratio = cx / cy
    with Image.open(image_path) as image:
        image_ratio = image.width / image.height
    left = top = right = bottom = 0
    if image_ratio > frame_ratio:
        excess = 1 - frame_ratio / image_ratio
        left = right = int(excess / 2 * 100000)
    elif image_ratio < frame_ratio:
        excess = 1 - image_ratio / frame_ratio
        top = bottom = int(excess / 2 * 100000)
    return {'l': str(left), 't': str(top), 'r': str(right), 'b': str(bottom)}


def replace_slot(slide, kind, container, blip, image_path: Path):
    image_part, r_id = slide.part.get_or_add_image_part(str(image_path))
    blip.set(EMBED, r_id)
    fill = container.find('.//' + qn('a:blipFill'))
    rect = cover_src_rect(container, image_path)
    if fill is not None and rect is not None:
        for old in fill.findall(qn('a:srcRect')):
            fill.remove(old)
        src_rect = fill.makeelement(qn('a:srcRect'), rect)
        position = list(fill).index(blip) + 1 if blip in list(fill) else 1
        fill.insert(position, src_rect)
    return image_part


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--pptx', required=True)
    parser.add_argument('--images', required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--page', default='', help='只换这一页（1 起）；留空=整册')
    args = parser.parse_args()

    images = [Path(part.strip()) for part in args.images.split(',') if part.strip()]
    if not images:
        raise SystemExit('至少给一张图')
    for image in images:
        if not image.is_file():
            raise SystemExit('图片不存在：' + str(image))

    source = Path(args.pptx)
    target = Path(args.out)
    if not source.is_file():
        raise SystemExit('文件不存在：' + str(source))
    target.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != target.resolve():
        shutil.copy2(source, target)

    deck = Presentation(str(target))
    only = int(args.page) if str(args.page).strip() else 0
    replaced = 0
    cursor = 0
    for index, slide in enumerate(deck.slides, start=1):
        if only and index != only:
            continue
        for kind, container, blip in slot_elements(slide):
            image_path = images[cursor % len(images)]
            cursor += 1
            replace_slot(slide, kind, container, blip, image_path)
            replaced += 1

    deck.save(str(target))
    print('输出:', target)
    print('替换图片槽: %d | 页数: %d | 图数: %d' % (replaced, len(deck.slides), len(images)))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
