# -*- coding: utf-8 -*-
"""效果14（立体图片轮播）参数化构建：N 张图 → N+1 页。

规则见 effects/014_图片轮播/recipe.md：
  P1：C 位叠 图1 + 图2；Pk：A=图k B=图k+1 C=图k+2 D=图k−1；越界形状移出画布
  标签只减不增：隐藏 >N 的编号，按 SLOT=33.87/N 重排；高亮块对准当前页；文字小副本隐藏
  图片按目标框做 cover 裁切（写 a:srcRect），避免变形

用法：python build_effect14_dynamic.py --images a.jpg,b.jpg,c.jpg --out out.pptx
"""
import argparse
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

LIB = Path(r'D:/DSH效果库/ppt-effects')
TEMPLATE = LIB / 'effects' / '014_图片轮播' / 'reference' / '效果14.pptx'
EMU = 360000.0
SLOT_TOTAL = 33.87
LABEL_Y = 1.4
HIDDEN_Y = -20.0
OFF_Y = -50.0
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def slot_of(shape) -> str:
    x = (shape.left or 0) / EMU
    if x > 30:
        return 'D'
    if x > 10:
        return 'A'
    if x > -1.5:
        return 'B'
    return 'C'


def set_picture(slide, shape, image_path: Path):
    image_part, r_id = slide.part.get_or_add_image_part(str(image_path))
    blip = shape._element.find('.//' + qn('a:blip'))
    if blip is None:
        return
    blip.set(qn('r:embed'), r_id)
    apply_cover_crop(shape, image_path)


def apply_cover_crop(shape, image_path: Path):
    """按框宽高比给图片写 a:srcRect（cover 裁切），不改形状几何。"""
    blip_fill = shape._element.find('.//' + qn('p:blipFill'))
    if blip_fill is None:
        return
    for old in blip_fill.findall(qn('a:srcRect')):
        blip_fill.remove(old)
    frame_ratio = (shape.width or 1) / (shape.height or 1)
    with Image.open(image_path) as image:
        image_ratio = image.width / image.height
    l = t = r = b = 0
    if image_ratio > frame_ratio:
        excess = 1 - frame_ratio / image_ratio
        l = int(excess / 2 * 100000)
        r = l
    elif image_ratio < frame_ratio:
        excess = 1 - image_ratio / frame_ratio
        t = int(excess / 2 * 100000)
        b = t
    src = blip_fill.makeelement(qn('a:srcRect'), {'l': str(l), 't': str(t), 'r': str(r), 'b': str(b)})
    blip = blip_fill.find(qn('a:blip'))
    position = list(blip_fill).index(blip) + 1 if blip is not None else 0
    blip_fill.insert(position, src)


def move(shape, y_cm: float):
    shape.top = int(y_cm * EMU)


def build(images, out_path: Path):
    count = len(images)
    if not 3 <= count <= 10:
        raise SystemExit('图片张数必须是 3–10（页数 = N+1）')
    prs = Presentation(str(TEMPLATE))

    keep = count + 1
    sld_id_list = prs.slides._sldIdLst
    for sld_id in list(sld_id_list)[keep:]:
        r_id = sld_id.rId
        sld_id_list.remove(sld_id)
        prs.part.drop_rel(r_id)
    slides = list(prs.slides)

    for page_no, slide in enumerate(slides, start=1):
        pics = [sh for sh in slide.shapes if sh._element.tag == qn('p:pic')]
        if page_no == 1:
            c_shapes = sorted([sh for sh in pics if slot_of(sh) == 'C'], key=lambda sh: sh.left)
            for order, shape in enumerate(c_shapes):
                index = order + 1
                if index <= count:
                    set_picture(slide, shape, images[index - 1])
                else:
                    move(shape, OFF_Y)
            continue
        k = page_no - 1
        for shape in pics:
            index = {'A': k, 'B': k + 1, 'C': k + 2, 'D': k - 1}[slot_of(shape)]
            if 1 <= index <= count:
                set_picture(slide, shape, images[index - 1])
            else:
                move(shape, OFF_Y)

    slot = SLOT_TOTAL / count
    for page_no, slide in enumerate(slides, start=1):
        by_number = {}
        for shape in slide.shapes:
            if not getattr(shape, 'has_text_frame', False):
                continue
            text = (shape.text_frame.text or '').strip()
            if len(text) == 2 and text.isdigit():
                by_number[int(text)] = shape
        for number, shape in by_number.items():
            if number > count:
                move(shape, HIDDEN_Y)
                continue
            center = slot * (number - 0.5)
            shape.left = int((center - shape.width / EMU / 2) * EMU)
            shape.top = int(LABEL_Y * EMU)
        for shape in slide.shapes:
            if shape.name.startswith('矩形: 圆角'):
                if page_no == 1:
                    shape.left = int(-6 * EMU)
                    shape.top = int(0.93 * EMU)
                else:
                    target = by_number.get(page_no - 1) or by_number.get(page_no)
                    if target is not None:
                        shape.left = int(target.left + target.width / 2 - shape.width / 2)
                        shape.top = int(target.top + target.height / 2 - shape.height / 2)
            if shape.name in ('TextBox 13', 'TextBox 15'):
                move(shape, -40.0)

    sys.path.insert(0, str(LIB / 'tools'))
    from ppt_morph import ensure_morph
    added, skipped = ensure_morph(prs)
    print('平滑：新增', added, '页，已有', skipped, '页')

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print('输出:', out_path)
    print('页数:', len(slides), '（N=%d）' % count)
    return out_path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--images', required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--preview', default='')
    args = parser.parse_args()
    images = [Path(part.strip()) for part in args.images.split(',') if part.strip()]
    for image in images:
        if not image.is_file():
            raise SystemExit('图片不存在：' + str(image))
    out = build(images, Path(args.out))
    if args.preview:
        import subprocess
        subprocess.run([sys.executable, str(LIB / 'tools' / 'render_preview.py'), str(out), args.preview, '1,3'], check=False)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
