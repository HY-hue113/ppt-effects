# -*- coding: utf-8 -*-
"""A 批：把「原生形状 + 局部图片」页里的图片改成可填充的图片占位

保留全部原生形状，只把 p:pic 替换成虚线占位框（内写「图片」），
用户可直接右键「设置形状格式 → 填充 → 图片」换图。

批次 B 页：p19、p23、p25、p27
"""
import os
import zipfile

from lxml import etree

LIB = r'<LIB_ROOT>\layout-tools'
SRC = os.path.join(LIB, 'reference', '版式库-01-06.pptx')
OUT = os.path.join(LIB, 'reference', '版式库-占位-B.pptx')
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525
PAGES_A = [19, 23, 25, 27]


def placeholder(sid, x, y, w, h, label='图片'):
    return ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s占位"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="F7F7F7"/></a:solidFill>'
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>'
            '<a:prstDash val="dash"/></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
            '<a:r><a:rPr lang="zh-CN" sz="1600"><a:solidFill><a:srgbClr val="8C8C8C"/>'
            '</a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
            % (P, A, sid, label, x * EMU, y * EMU, w * EMU, h * EMU, label))


def main():
    z = zipfile.ZipFile(SRC)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]

    total = 0
    for idx in PAGES_A:
        part = parts[idx - 1]
        root = etree.fromstring(items[part])
        spTree = root.find('.//{%s}spTree' % P)
        n = 0
        for pic in list(root.iter('{%s}pic' % P)):
            xf = pic.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
            if off is None or ext is None:
                continue
            x = int(off.get('x')) // EMU
            y = int(off.get('y')) // EMU
            w = int(ext.get('cx')) // EMU
            h = int(ext.get('cy')) // EMU
            if w < 40 or h < 30:          # 小装饰图跳过
                continue
            nv = pic.find('.//{%s}cNvPr' % P)
            sid = int(nv.get('id')) if nv is not None and nv.get('id', '').isdigit() else 900 + n
            parent = pic.getparent()
            parent.insert(list(parent).index(pic), etree.fromstring(placeholder(sid, x, y, w, h)))
            parent.remove(pic)
            n += 1
        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        total += n
        print('p%-2d 图片转占位 %d 个' % (idx, n))

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    print('B 批占位版完成，共 %d 个图片位 → %s' % (total, OUT))


if __name__ == '__main__':
    main()
