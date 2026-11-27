# -*- coding: utf-8 -*-
"""效果11（中式滑转 · 多图版）参数化入口：a 张图 → a+1 页。

直接复用效果自带的原始脚本 `effects/011_中式滑转/tools/eff11_multi.py`
（"复制 p2/p3 骨架 + 改名 + 填图/上色 + 末页奇偶镜像"），只把脚本里的
`IMGS` / `SRC` / `OUT` 三个常量替换成本次参数后执行，避免重写带来的偏差。

用法：
  python build_effect11_dynamic.py --images a.jpg,b.jpg,c.jpg --out 效果11-3图.pptx
"""
import argparse
import re
import re
import runpy
import zipfile
import sys
import tempfile

from lxml import etree
from pathlib import Path

LIB = Path(r'D:/DSH效果库/ppt-effects')
LEGACY = LIB / 'effects' / '011_中式滑转' / 'tools' / 'eff11_multi.py'
TEMPLATE = LIB / 'effects' / '011_中式滑转' / 'reference' / '效果11-多图版.pptx'


def run(images, out_path: Path):
    count = len(images)
    if not 2 <= count <= 9:
        raise SystemExit('图片张数必须是 2–9（页数 = N+1）')
    source = LEGACY.read_text(encoding='utf-8')

    # 替换 IMGS 列表
    imgs_block = 'IMGS = [\n' + ''.join("    r'%s',\n" % str(image).replace('\\', '/') for image in images) + ']\n'
    source, replaced_images = re.subn(r'IMGS = \[.*?\n\]\n', imgs_block, source, count=1, flags=re.S)
    if replaced_images != 1:
        raise SystemExit('legacy 脚本里找不到 IMGS 列表')

    # 替换 SRC / OUT
    source, replaced_src = re.subn(r"^SRC = r'.*?'$", "SRC = r'%s'" % str(TEMPLATE).replace('\\', '/'), source, count=1, flags=re.M)
    source, replaced_out = re.subn(r"^OUT = r'.*?'$", "OUT = r'%s'" % str(out_path).replace('\\', '/'), source, count=1, flags=re.M)
    if replaced_src != 1 or replaced_out != 1:
        raise SystemExit('legacy 脚本里找不到 SRC / OUT')

    with tempfile.NamedTemporaryFile('w', suffix='.py', delete=False, encoding='utf-8') as handle:
        handle.write(source)
        temp_path = handle.name
    try:
        runpy.run_path(temp_path, run_name='__main__')
    finally:
        Path(temp_path).unlink(missing_ok=True)
    prune_unused_slides(out_path)
    sys.path.insert(0, str(LIB / 'tools'))
    from ppt_morph import ensure_morph
    from pptx import Presentation

    prs = Presentation(str(out_path))
    added, skipped = ensure_morph(prs)
    temp = out_path.with_name(out_path.stem + '-morph.pptx')
    prs.save(str(temp))
    temp.replace(out_path)
    print('平滑：新增', added, '页，已有', skipped, '页')
    print('页数:', count + 1, '（N=%d）' % count)
    return out_path



P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'


def prune_unused_slides(path: Path):
    """原始脚本会留下模板里未被引用的 slide 部件；这里按 sldIdLst 清单删掉，保持包干净。"""
    with zipfile.ZipFile(path) as archive:
        items = {name: archive.read(name) for name in archive.namelist()}
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    sld_lst = pres.find('{%s}sldIdLst' % P_NS)
    if sld_lst is None:
        return
    rel_target = {rel.get('Id'): rel.get('Target') for rel in rels}
    used = set()
    for sld_id in sld_lst:
        target = rel_target.get(sld_id.get('{%s}id' % R_NS)) or ''
        match = re.fullmatch(r'slides/slide(\d+)\.xml', target)
        if match:
            used.add(int(match.group(1)))
    types = etree.fromstring(items['[Content_Types].xml'])
    for key in list(items):
        match = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', key)
        if not match:
            continue
        number = int(match.group(1))
        if number in used:
            continue
        items.pop(key, None)
        items.pop('ppt/slides/_rels/slide%d.xml.rels' % number, None)
    for element in list(types):
        part = element.get('PartName') or ''
        match = re.fullmatch(r'/ppt/slides/slide(\d+)\.xml', part)
        if match and int(match.group(1)) not in used:
            types.remove(element)
    items['[Content_Types].xml'] = etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for key, data in items.items():
            archive.writestr(key, data)
    print('已清理未引用 slide 部件，剩余页数:', len(used))
def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--images', required=True)
    parser.add_argument('--out', required=True)
    args = parser.parse_args()
    images = [Path(part.strip()) for part in args.images.split(',') if part.strip()]
    for image in images:
        if not image.is_file():
            raise SystemExit('图片不存在：' + str(image))
    run(images, Path(args.out))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
