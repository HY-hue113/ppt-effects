# -*- coding: utf-8 -*-
"""把 typeset 引擎跑一遍版式库 22 页，检查副作用

只调字号（不改几何、不重建 run），输出对比稿 + 变更清单。
"""
import json
import os
import sys
import zipfile

from lxml import etree

sys.path.insert(0, r'<LIB_ROOT>\tools')
import typeset  # noqa: E402

LIB = r'<LIB_ROOT>\layout-tools'
SRC = os.path.join(LIB, 'reference', '版式库-参考件-22页.pptx')
OUT = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\out\typeset\layout-typeset.pptx'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def cur_size(sp):
    for r in sp.iter('{%s}r' % A):
        rPr = r.find('{%s}rPr' % A)
        if rPr is not None and rPr.get('sz'):
            return int(rPr.get('sz')) // 100
    return None


def main():
    with zipfile.ZipFile(SRC) as z:
        items = {n: z.read(n) for n in z.namelist()}
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]

    changes, overflowed = [], []
    for idx, part in enumerate(parts, 1):
        root = etree.fromstring(items[part])
        touched = False
        for sp in root.iter('{%s}sp' % P):
            texts = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if not texts:
                continue
            box = typeset.shape_box(sp)
            if not box:
                continue
            _, _, w, h = box
            if h < 40:      # 矮标签框：原稿即溢出，不参与排版计算
                continue
            before = cur_size(sp)
            padx = min(32, max(6, w * 0.12))
            pady = min(24, max(4, h * 0.15))
            ml = 1 if h < 50 else 3
            r = typeset.fit_block(texts, max(w - padx * 2, 12), max(h - pady * 2, 10),
                                  max_pt=before or 28, min_pt=20, max_lines=ml)
            if r['overflow']:
                overflowed.append({'page': idx, 'text': texts[:18], 'need_h': r['need_h'], 'box_h': h})
            if before is not None and r['pt'] != before:
                changes.append({'page': idx, 'text': texts[:18], 'from': before, 'to': r['pt'],
                                'lines': len(r['lines'])})
                for run in sp.iter('{%s}r' % A):
                    rPr = run.find('{%s}rPr' % A)
                    if rPr is not None:
                        rPr.set('sz', str(int(r['pt'] * 100)))
                touched = True
        if touched:
            items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)

    print('页数', len(parts))
    print('字号变化', len(changes), '处')
    for c in changes[:15]:
        print('  p%-2d %-18s %d → %dpt (%d 行)' % (c['page'], c['text'], c['from'], c['to'], c['lines']))
    print('溢出（未处理）', len(overflowed), '处')
    for o in overflowed[:10]:
        print('  p%-2d %-18s 需高 %d / 框高 %d' % (o['page'], o['text'], o['need_h'], o['box_h']))
    json.dump({'changes': changes, 'overflowed': overflowed},
              open(os.path.join(os.path.dirname(OUT), 'layout-typeset-report.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
