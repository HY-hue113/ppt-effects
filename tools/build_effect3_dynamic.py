# -*- coding: utf-8 -*-
"""效果3（哲学字动 / love字动）参数化构建：给 N 张图，自动出 2N+2 页。

规则来自 effects/003_哲学字动/spec.json：
  - 页 = 2 + 2N（2 页开场文字 + N 组「照片页 + 文字页」）
  - 参考件 5 个照片槽（按画幅）：1.76 / 1.58 / 1.32 / 1.89 / 1.68
  - N 取 1–5；少于 5 张时按槽位画幅顺序取前 N 个槽

做法：模板移植——保留参考件的 slide XML（几何/动画/切换/音频全部原样），
只删掉多余页、把新图按目标槽位的像素尺寸（cover 裁切）替换进去。

用法：
  python build_effect3_dynamic.py --images a.jpg,b.jpg,c.jpg,d.jpg --out 效果3-4图.pptx
  python build_effect3_dynamic.py --images a.jpg --out 效果3-1图.pptx --preview out_dir
"""
import argparse
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

from PIL import Image
from pptx import Presentation

LIB = Path(r'D:/DSH效果库/ppt-effects')
TEMPLATE = LIB / 'effects' / '003_哲学字动' / 'reference' / '效果3.pptx'

# 参考件照片页（第 3/5/7/9/11 页）→ 媒体文件 + 画幅
SLOTS = [
    (3, 'ppt/media/image2.jpeg', 1.76),
    (5, 'ppt/media/image3.jpg', 1.58),
    (7, 'ppt/media/image4.jpeg', 1.32),
    (9, 'ppt/media/image5.jpeg', 1.89),
    (11, 'ppt/media/image6.jpeg', 1.68),
]
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'


def cover_resize(image: Image.Image, target: tuple) -> Image.Image:
    """按目标尺寸裁切填充（cover），再 LANCZOS 缩放到目标像素。"""
    tw, th = target
    if tw <= 0 or th <= 0:
        return image
    src_ratio = image.width / image.height
    tgt_ratio = tw / th
    if src_ratio > tgt_ratio:
        new_w = int(image.height * tgt_ratio)
        left = (image.width - new_w) // 2
        image = image.crop((left, 0, left + new_w, image.height))
    else:
        new_h = int(image.width / tgt_ratio)
        top = (image.height - new_h) // 2
        image = image.crop((0, top, image.width, top + new_h))
    return image.resize((tw, th), Image.LANCZOS)


def kept_pages(count: int):
    """N 张图 → 保留的页码列表（1 基）。"""
    pages = [1, 2]
    for index in range(count):
        pages.extend([3 + 2 * index, 4 + 2 * index])
    return pages


def build(images, out_path: Path):
    if not 1 <= len(images) <= 5:
        raise SystemExit('图片张数必须是 1–5（页数 = 2N+2）')
    with zipfile.ZipFile(TEMPLATE) as source:
        items = {name: source.read(name) for name in source.namelist()}
        infos = {info.filename: info for info in source.infolist()}

    kept = kept_pages(len(images))
    total = 12
    removed = [page for page in range(1, total + 1) if page not in kept]

    presentation_rels = items['ppt/_rels/presentation.xml.rels'].decode('utf-8')
    rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', presentation_rels))
    slide_rid = {}
    for rid, target in rid_target.items():
        if target.startswith('slides/slide'):
            number = int(re.search(r'slide(\d+)\.xml', target).group(1))
            slide_rid[number] = rid

    # 1) presentation.xml 去掉被删页的 sldId
    xml = items['ppt/presentation.xml'].decode('utf-8')
    for page in removed:
        rid = slide_rid.get(page)
        if rid:
            xml = re.sub(r'<p:sldId [^>]*r:id="%s"\s*/>' % re.escape(rid), '', xml)
    items['ppt/presentation.xml'] = xml.encode('utf-8')

    # 2) presentation rels 去掉被删页
    for page in removed:
        rid = slide_rid.get(page)
        if rid:
            presentation_rels = re.sub(r'<Relationship [^>]*Id="%s"[^>]*/>' % re.escape(rid), '', presentation_rels)
    items['ppt/_rels/presentation.xml.rels'] = presentation_rels.encode('utf-8')

    # 3) 删 slide 部件 + rels + Content_Types Override
    content_types = items['[Content_Types].xml'].decode('utf-8')
    for page in removed:
        for name in ('ppt/slides/slide%d.xml' % page, 'ppt/slides/_rels/slide%d.xml.rels' % page):
            items.pop(name, None)
        content_types = re.sub(r'<Override PartName="/ppt/slides/slide%d\.xml"[^>]*/>' % page, '', content_types)
    items['[Content_Types].xml'] = content_types.encode('utf-8')

    # 4) 替换照片：按槽位顺序取前 N 个
    for slot_index, image_path in enumerate(images):
        page, media_name, aspect = SLOTS[slot_index]
        if page not in kept:
            raise SystemExit('内部错误：槽位页 %d 不在保留页里' % page)
        with Image.open(image_path) as source_image:
            source_image = source_image.convert('RGB')
            with Image.open(BytesIO(items[media_name])) as target:
                target_size = target.size
            fitted = cover_resize(source_image, target_size)
            actual = source_image.width / source_image.height
            if abs(actual - aspect) / aspect > 0.12:
                print('  ! 第 %d 张图 画幅 %.2f 与槽位 %.2f 差得多，已 cover 裁切（会裁掉一部分）'
                      % (slot_index + 1, actual, aspect))
            buffer = BytesIO()
            suffix = Path(media_name).suffix.lower()
            if suffix in ('.jpg', '.jpeg'):
                fitted.save(buffer, format='JPEG', quality=92)
            else:
                fitted.save(buffer, format='PNG')
            items[media_name] = buffer.getvalue()

    # 5) 只保留仍被引用的媒体（被删页独占的媒体一并删掉）
    referenced = set()
    for name, data in items.items():
        if not name.endswith('.rels'):
            continue
        text = data.decode('utf-8')
        for target in re.findall(r'Target="([^"]+)"', text):
            if 'media/' not in target:
                continue
            base = name.replace('_rels/', '').rsplit('/', 1)[0]
            resolved = str(Path(base) / target).replace('\\', '/')
            parts = []
            for piece in resolved.split('/'):
                if piece == '..':
                    parts.pop()
                elif piece and piece != '.':
                    parts.append(piece)
            referenced.add('/'.join(parts))
    for name in list(items):
        if name.startswith('ppt/media/') and name not in referenced:
            items.pop(name, None)
            ct = items['[Content_Types].xml'].decode('utf-8')
            ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(name), '', ct)
            items['[Content_Types].xml'] = ct.encode('utf-8')

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as target_zip:
        for name, data in items.items():
            info = infos.get(name)
            target_zip.writestr(info or name, data)
    adjust_letters_and_morph(out_path, len(images))
    print('输出:', out_path)
    print('页数:', 2 + 2 * len(images), '（N=%d）' % len(images))
    return out_path



def adjust_letters_and_morph(path: Path, count: int):
    """字母对象数 = 图数 N；紧凑页均分、图片页只留当前一个字母，其余移出画布；再补平滑。"""
    sys.path.insert(0, str(LIB / 'tools'))
    from ppt_morph import ensure_morph
    prs = Presentation(str(path))
    width = prs.slide_width
    for page_index, slide in enumerate(prs.slides, start=1):
        letters = []
        labels = []
        for shape in slide.shapes:
            if not getattr(shape, 'has_text_frame', False):
                continue
            value = shape.text_frame.text.strip()
            if len(value) == 1 and value.isalpha():
                letters.append(shape)
            elif value:
                labels.append(shape)
        if not letters:
            continue
        letters.sort(key=lambda item: item.left or 0)
        used = set()
        pairs = []
        for letter in letters:
            center = (letter.left or 0) + (letter.width or 0) / 2
            best = None
            best_delta = 1e18
            for label in labels:
                if id(label._element) in used:
                    continue
                delta = abs(((label.left or 0) + (label.width or 0) / 2) - center)
                if delta < best_delta:
                    best_delta = delta
                    best = label
            if best is not None:
                used.add(id(best._element))
            pairs.append((letter, best))
        # 超出 N 的字母/标签移出画布
        for letter, label in pairs[count:]:
            letter.top = int(-50 * 360000)
            if label is not None:
                label.top = int(-50 * 360000)
        pairs = pairs[:count]
        if not pairs:
            continue
        # 图片页（第 3/5/7… 页）只留当前字母，其余移出画布；其他页均分
        photo_page = page_index >= 3 and page_index % 2 == 1
        if photo_page:
            current = (page_index - 1) // 2  # 1..N
            if current > len(pairs):
                current = len(pairs)
            for index, (letter, label) in enumerate(pairs, start=1):
                if index != current:
                    letter.top = int(-50 * 360000)
                    if label is not None:
                        label.top = int(-50 * 360000)
                    continue
                center = width * (0.06 if index % 2 == 1 else 0.68)
                letter.left = int(center - (letter.width or 0) / 2)
                if label is not None:
                    label.left = int(center - (label.width or 0) / 2)
            continue
        for index, (letter, label) in enumerate(pairs):
            center = width * (1 / 8 + (6 / 8) * index / max(1, len(pairs) - 1))
            letter.left = int(center - (letter.width or 0) / 2)
            if label is not None:
                label.left = int(center - (label.width or 0) / 2)
    added, skipped = ensure_morph(prs)
    temp = path.with_name(path.stem + '-tmp.pptx')
    prs.save(str(temp))
    Presentation(str(temp)).save(str(path))
    temp.unlink(missing_ok=True)
    print('字母对象:', count, '| 平滑：新增', added, '页，已有', skipped, '页')
def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--images', required=True, help='逗号分隔的图片路径，1–5 张')
    parser.add_argument('--out', required=True)
    parser.add_argument('--preview', default='', help='顺带渲染首页到这个目录')
    args = parser.parse_args()
    images = [part.strip() for part in args.images.split(',') if part.strip()]
    for image in images:
        if not Path(image).is_file():
            raise SystemExit('图片不存在：' + image)
    out = build(images, Path(args.out))
    if args.preview:
        import subprocess
        subprocess.run([sys.executable, str(LIB / 'tools' / 'render_preview.py'), str(out), args.preview, '1'], check=False)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
