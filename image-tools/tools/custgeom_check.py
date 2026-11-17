# -*- coding: utf-8 -*-
"""custGeom 合法性校验（防止 PowerPoint 报"需要修复"）

三条硬要求：
  1. <a:path ... fill="norm">        —— 否则填充被当成"无"
  2. spPr 子元素顺序 xfrm → (custGeom) → fill → ln
  3. 路径首尾闭合、坐标 ∈ [0, 100000]

用法：
    from custgeom_check import check_custgeom, check_sppr_order
    check_custgeom(geom_xml)          # 抛异常即不合法
"""

from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

_SKIP = ('avLst', 'gdLst', 'ahLst', 'cxnLst', 'rect', 'pathLst')


class CustGeomError(ValueError):
    pass


def check_custgeom(geom_xml: str, tol: float = 1.0):
    """校验 custGeom：fill="norm"、坐标范围、路径闭合"""
    root = etree.fromstring(geom_xml)
    paths = root.findall('.//{%s}path' % A)
    if not paths:
        raise CustGeomError('没有 <a:path>')
    for i, path in enumerate(paths):
        if path.get('fill') != 'norm':
            raise CustGeomError('第 %d 条 path 缺少 fill="norm"（填充会失效）' % (i + 1))
        for pt in path.iter('{%s}pt' % A):
            x, y = int(pt.get('x')), int(pt.get('y'))
            if not (0 <= x <= 100000 and 0 <= y <= 100000):
                raise CustGeomError('坐标越界: (%d, %d)' % (x, y))
        if path.find('{%s}close' % A) is None:
            raise CustGeomError('第 %d 条 path 未闭合（缺 <a:close/>）' % (i + 1))
    return True


def check_sppr_order(sp_xml: str):
    """校验 spPr 子元素顺序：xfrm → geom → 填充 → ln → 效果"""
    root = etree.fromstring(sp_xml)
    spPr = root.find('{%s}spPr' % P)
    if spPr is None:
        raise CustGeomError('没有 spPr')
    order = [etree.QName(c).localname for c in spPr]
    geom = [n for n in order if n in ('prstGeom', 'custGeom')]
    if not geom:
        raise CustGeomError('缺少几何')
    gi = order.index(geom[0])
    fills = [n for n in order if n in ('solidFill', 'gradFill', 'blipFill', 'noFill', 'pattFill')]
    if fills:
        fi = order.index(fills[0])
        if fi < gi:
            raise CustGeomError('填充出现在几何之前（顺序应为 xfrm → geom → fill → ln）')
    if 'ln' in order and fills and order.index('ln') < order.index(fills[0]):
        raise CustGeomError('描边出现在填充之前')
    return order


if __name__ == '__main__':
    demo = ('<a:custGeom xmlns:a="%s"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="100000" b="100000"/>'
            '<a:pathLst><a:path w="100000" h="100000" fill="norm">'
            '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
            '<a:lnTo><a:pt x="100000" y="0"/></a:lnTo>'
            '<a:lnTo><a:pt x="100000" y="100000"/></a:lnTo>'
            '<a:close/></a:path></a:pathLst></a:custGeom>' % A)
    print('demo custGeom:', check_custgeom(demo))
