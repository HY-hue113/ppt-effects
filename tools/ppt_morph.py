# -*- coding: utf-8 -*-
"""平滑切换（morph）工具。

给 pptx 每页补「平滑（morph, 按对象）」切换，结构与参考件一致：
  <mc:AlternateContent><mc:Choice Requires="p159">
    <p:transition spd="slow"><p159:morph option="byObject"/></p:transition>
  </mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>
  </mc:AlternateContent>

模块：from ppt_morph import ensure_morph; ensure_morph(prs, force=False)
命令行：python ppt_morph.py --pptx in.pptx --out out.pptx [--force]

--force：已有非平滑切换的页也改成平滑；默认只补没有切换的页。
"""
import argparse
from pathlib import Path

from lxml import etree
from pptx import Presentation

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P159_NS = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'


def _localname(node):
    return etree.QName(node).localname if node is not None else ''


def _existing_transitions(root):
    return root.findall('.//{%s}transition' % P_NS)


def the_morph_block():
    block = etree.Element('{%s}AlternateContent' % MC_NS, nsmap={'mc': MC_NS, 'p159': P159_NS})
    choice = etree.SubElement(block, '{%s}Choice' % MC_NS)
    choice.set('Requires', 'p159')
    transition = etree.SubElement(choice, '{%s}transition' % P_NS)
    transition.set('spd', 'slow')
    morph = etree.SubElement(transition, '{%s}morph' % P159_NS)
    morph.set('option', 'byObject')
    fallback = etree.SubElement(block, '{%s}Fallback' % MC_NS)
    fallback_transition = etree.SubElement(fallback, '{%s}transition' % P_NS)
    fallback_transition.set('spd', 'slow')
    etree.SubElement(fallback_transition, '{%s}fade' % P_NS)
    return block


def ensure_morph(prs, force=False):
    added = 0
    skipped = 0
    for slide in prs.slides:
        root = slide._element
        transitions = _existing_transitions(root)
        has_morph = any(item.find('{%s}morph' % P159_NS) is not None for item in transitions)
        if has_morph and not force:
            skipped += 1
            continue
        for item in transitions:
            carrier = item
            while carrier.getparent() is not None and _localname(carrier.getparent()) != 'sld':
                carrier = carrier.getparent()
            parent = carrier.getparent()
            if parent is not None:
                parent.remove(carrier)
        csld = root.find('{%s}cSld' % P_NS)
        block = the_morph_block()
        if csld is not None:
            csld.addnext(block)
        else:
            root.insert(0, block)
        added += 1
    return added, skipped


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--pptx', required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args()
    prs = Presentation(args.pptx)
    added, skipped = ensure_morph(prs, force=args.force)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print('输出:', out, '| 补平滑:', added, '页 | 已有/跳过:', skipped, '页')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
