# -*- coding: utf-8 -*-
"""模板后处理：统一前五页表单字段文案 + 图表占位检查

用法：python post_labels.py <pptx>
"""
import re
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

CANON = [
    ("指导教师", "指导教师：XXX 教授"),
    ("指导老师", "指导教师：XXX 教授"),
    ("课题小组", "课题小组：XXX"),
    ("专业", "专业：xxx"),
    ("答辩人", "答辩人：xxx"),
    ("研究方向", "研究方向：xxx"),
    ("导师", "导师：xxx"),
    ("答辩日期", "日期：xxxx/xx/xx"),
    ("日期", "日期：xxxx/xx/xx"),
]


def set_text(sp, text):
    txBody = sp.find('{%s}txBody' % P)
    if txBody is None:
        return False
    paras = txBody.findall('{%s}p' % A)
    if not paras:
        return False
    first = paras[0]
    runs = first.findall('{%s}r' % A)
    rPr = runs[0].find('{%s}rPr' % A) if runs else None
    for p in paras[1:]:
        txBody.remove(p)
    for r in first.findall('{%s}r' % A):
        first.remove(r)
    r = etree.SubElement(first, '{%s}r' % A)
    if rPr is not None:
        r.append(rPr)
    t = etree.SubElement(r, '{%s}t' % A)
    t.text = text
    return True


def main(path):
    z = zipfile.ZipFile(path)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    slides = sorted([n for n in items if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)],
                    key=lambda s: int(re.findall(r'\d+', s)[0]))
    fixed = 0
    for idx, n in enumerate(slides, 1):
        if idx > 5:
            break
        root = etree.fromstring(items[n])
        for sp in root.iter('{%s}sp' % P):
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if not txt:
                continue
            for key, canon in CANON:
                if key in txt:
                    if set_text(sp, canon):
                        fixed += 1
                    break
        items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # NAV_FIX：顶部导航标签（y<140）上的「正文」改为「标题」
    nav = 0
    for n in slides:
        root = etree.fromstring(items[n])
        changed = False
        for sp in root.iter('{%s}sp' % P):
            txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
            if txt != '正文':
                continue
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off = xf.find('{%s}off' % A)
            if off is None:
                continue
            if int(off.get('y')) // 9525 < 140:
                if set_text(sp, "标题"):
                    nav += 1
                    changed = True
        if changed:
            items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    print('导航标签改标题:', nav)

    # 统计图表
    charts = 0
    for n in slides:
        s = items[n].decode('utf-8', 'ignore')
        charts += s.count('<c:chart')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return fixed, charts


if __name__ == '__main__':
    f, c = main(sys.argv[1])
    print('前五页字段统一 %d 处；图表引用 %d 个' % (f, c))
