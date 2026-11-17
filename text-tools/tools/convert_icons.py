# -*- coding: utf-8 -*-
"""把小图标图片（p:pic）替换成可编辑矢量形状，保证"每个都能改、都能换色"

规则：
    图片 → 同尺寸圆角矩形（默认主色填充、半透明），可改颜色/改形状
    0×0 的图片直接删除（空对象）
"""
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
EMU = 9525


def convert(path, color='C00000', alpha=45000):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    n_rect = n_del = 0
    for name in [x for x in items if x.startswith('ppt/slides/slide') and x.endswith('.xml')]:
        root = etree.fromstring(items[name])
        changed = False
        for pic in list(root.iter('{%s}pic' % P)):
            xf = pic.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
            if off is None or ext is None:
                continue
            w, h = int(ext.get('cx')), int(ext.get('cy'))
            parent = pic.getparent()
            if w == 0 or h == 0:
                parent.remove(pic)
                n_del += 1
                changed = True
                continue
            nv = pic.find('.//{%s}cNvPr' % P)
            sid = nv.get('id') if nv is not None else '900'
            xml = (
                '<p:sp xmlns:p="%s" xmlns:a="%s">'
                '<p:nvSpPr><p:cNvPr id="%s" name="图标位"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
                '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>'
                '<a:solidFill><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:solidFill>'
                '<a:ln w="6350"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
                '</p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>'
                '</p:sp>' % (P, A, sid, off.get('x'), off.get('y'), w, h, color, alpha, color))
            parent.insert(list(parent).index(pic), etree.fromstring(xml))
            parent.remove(pic)
            n_rect += 1
            changed = True
        if changed:
            items[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return n_rect, n_del


if __name__ == '__main__':
    for p in sys.argv[1:]:
        color = 'C00000' if '标题' in p or '文框' in p else '1843CE'
        a, b = convert(p, color=color)
        print('%s → 图标转形状 %d 个，删除空图 %d 个' % (p.split('\\')[-1], a, b))
