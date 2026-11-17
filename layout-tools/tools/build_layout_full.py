# -*- coding: utf-8 -*-
"""完整版式重绘：一次生成 版式库-重绘.pptx

1. 图片页（p5/11/15/19/21/23/25/26/27/30）：图片 → 可填充占位，大图沉底
2. 整页图片页（p6/7/8/10/9）：原生形状重画
3. 其余页保持原样
"""
import os
import sys
import zipfile

import numpy as np
from PIL import Image
from lxml import etree

LIB = r'<LIB_ROOT>\layout-tools'
SRC = os.path.join(LIB, 'reference', '版式库-01-06.pptx')
OUT = os.path.join(LIB, 'reference', '版式库-重绘.pptx')
IMG8 = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\out\layouts\p8.png'

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525
SW, SH = 1280, 720
AREA = SW * SH

PIC_PAGES = [5, 11, 15, 19, 21, 23, 25, 26, 27, 30]
DRAW_PAGES = [6, 7, 8, 9, 10]


# ---------- 基础图元 ----------
def sp_rect(sid, name, x, y, w, h, fill, geom='rect', adj=0, line=None, dash=False, rot=0):
    av = '<a:gd name="adj" fmla="val %d"/>' % adj if geom == 'roundRect' else ''
    ln = ('<a:ln w="12700"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>'
          % (line, '<a:prstDash val="dash"/>' if dash else '')) if line else '<a:ln><a:noFill/></a:ln>'
    return ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm%s><a:off x="%d" y="%d"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="%s"><a:avLst>%s</a:avLst>'
            '</a:prstGeom>%s%s</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'
            % (P, A, sid, name, (' rot="%d"' % rot) if rot else '', x * EMU, y * EMU,
               w * EMU, h * EMU, geom, av,
               ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill) if fill else '<a:noFill/>', ln))


def sp_poly(sid, name, pts, fill):
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    x0, y0 = min(xs), min(ys)
    w, h = max(max(xs) - x0, 1), max(max(ys) - y0, 1)

    def pt(p):
        return '<a:pt x="%d" y="%d"/>' % (int((p[0] - x0) / w * 100000), int((p[1] - y0) / h * 100000))

    cmds = ['<a:moveTo>%s</a:moveTo>' % pt(pts[0])] + ['<a:lnTo>%s</a:lnTo>' % pt(p) for p in pts[1:]] + ['<a:close/>']
    geom = ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="100000" b="100000"/>'
            '<a:pathLst><a:path w="100000" h="100000" fill="norm">%s</a:path></a:pathLst></a:custGeom>'
            % ''.join(cmds))
    return ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="%s"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm>%s<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'
            % (P, A, sid, name, x0 * EMU, y0 * EMU, w * EMU, h * EMU, geom, fill))


def sp_text(sid, x, y, w, h, content, size=16, color='595959'):
    return ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="T%d"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="%d"><a:solidFill>'
            '<a:srgbClr val="%s"/></a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p>'
            '</p:txBody></p:sp>'
            % (P, A, sid, sid, x * EMU, y * EMU, w * EMU, h * EMU, size * 100, color, content))


# ---------- 各页绘制 ----------
def draw_p6(t):
    boxes = [(40, 20, 400, 145, 'D9D9D9'), (40, 182, 560, 165, 'B7B2B2'),
             (40, 360, 270, 340, 'D9D9D9'), (322, 360, 278, 340, 'D9D9D9'),
             (462, 20, 380, 145, 'D9D9D9'), (462, 360, 380, 340, 'D9D9D9'),
             (860, 20, 380, 680, 'D9D9D9')]
    for i, (x, y, w, h, c) in enumerate(boxes):
        t.append(etree.fromstring(sp_rect(10 + i, '色块%d' % (i + 1), x, y, w, h, c)))
    t.append(etree.fromstring(sp_text(60, 40, 182, 560, 165, '图片', 22)))
    t.append(etree.fromstring(sp_text(61, 860, 20, 380, 680, '图片', 22)))
    t.append(etree.fromstring(sp_text(62, 40, 360, 270, 340, '文本', 20)))
    t.append(etree.fromstring(sp_text(63, 322, 360, 278, 340, '文本', 20)))


def draw_p7(t):
    t.append(etree.fromstring(sp_rect(100, '左列', 20, 20, 330, 680, 'D9D9D9')))
    t.append(etree.fromstring(sp_rect(101, '中列', 365, 0, 350, 720, 'B7B2B2')))
    t.append(etree.fromstring(sp_rect(102, '右列', 730, 20, 330, 680, 'D9D9D9')))
    for i, (cx, cy) in enumerate([(30, 150), (390, 340), (745, 70)]):
        t.append(etree.fromstring(sp_rect(110 + i, '圆%d' % (i + 1), cx, cy, 300, 300, 'D2D2D2', 'ellipse')))
        t.append(etree.fromstring(sp_text(120 + i, cx, cy, 300, 300, '图片', 18)))


def draw_p10(t):
    for i in range(4):
        t.append(etree.fromstring(sp_rect(200 + i, '长条%d' % (i + 1), -10 + i * 232, 15, 300, 160,
                                          ['DCDCDC', 'D2D2D2', 'C8C8C8', 'BEBEBE'][i], 'roundRect', 50000)))
    for i in range(4):
        cx = 40 + i * 305
        t.append(etree.fromstring(sp_rect(210 + i, '点%d' % (i + 1), cx - 5, 195, 12, 12, '7F7F7F', 'ellipse')))
        t.append(etree.fromstring(sp_rect(220 + i, '竖线%d' % (i + 1), cx, 205, 2, 500, 'D0D0D0')))
        t.append(etree.fromstring(sp_rect(230 + i, '小条1_%d' % (i + 1), cx + 30, 240, 110, 36, 'B7B2B2')))
        t.append(etree.fromstring(sp_rect(240 + i, '小条2_%d' % (i + 1), cx + 30, 282, 170, 36, 'B7B2B2')))
        t.append(etree.fromstring(sp_rect(250 + i, '图位%d' % (i + 1), cx + 30, 330, 240, 150, 'EFEFEF')))
        t.append(etree.fromstring(sp_text(260 + i, cx + 30, 330, 240, 150, '图片', 16)))


def fit_circle(pts):
    xs = np.array([p[0] for p in pts])
    ys = np.array([p[1] for p in pts])
    m = np.c_[xs, ys, np.ones(len(xs))]
    b = -(xs ** 2 + ys ** 2)
    D, E, F = np.linalg.lstsq(m, b, rcond=None)[0]
    cx, cy = -D / 2, -E / 2
    return cx, cy, np.sqrt(cx ** 2 + cy ** 2 - F)


def draw_p8(t):
    im = Image.open(IMG8).convert('L')
    W, H = im.size
    a = np.asarray(im).astype(int)
    sx, sy = W / SW, H / SH
    left, right = [], []
    for y in range(0, H, 4):
        idx = np.where((a[y] >= 185) & (a[y] <= 218))[0]
        for x in idx:
            X = x / sx
            if 330 < X < 640:
                left.append((X, y / sy))
                break
        for x in idx[::-1]:
            X = x / sx
            if 640 < X < 950:
                right.append((X, y / sy))
                break
    for pts, sid, nm in ((left, 300, '左弧'), (right, 301, '右弧')):
        if len(pts) > 10:
            cx, cy, r = fit_circle(pts)
            t.append(etree.fromstring(sp_rect(sid, nm, int(cx - r), int(cy - r),
                                              int(2 * r), int(2 * r), None, 'ellipse', 0, 'C9C9C9')))
    t.append(etree.fromstring(sp_rect(310, '中心图片位', 490, 185, 300, 300, 'F2F2F2', 'ellipse', 0, 'BFBFBF', True)))
    t.append(etree.fromstring(sp_text(311, 490, 185, 300, 300, '图片', 18)))
    for i, (cx, cy) in enumerate([(315, 150), (876, 150), (315, 490), (876, 490)]):
        t.append(etree.fromstring(sp_rect(320 + i, '小圆%d' % (i + 1), cx, cy, 90, 90, 'D9D9D9', 'ellipse')))
    blocks = [(60, 152, 250, 100), (970, 152, 250, 100), (60, 500, 250, 100), (970, 500, 250, 100)]
    for i, (bx, by, bw, bh) in enumerate(blocks):
        barx = bx + 155 if bx < 600 else bx - 45
        t.append(etree.fromstring(sp_rect(340 + i * 3, '深条%d' % (i + 1), barx, by - 42, 100, 20, 'A6A6A6')))
        t.append(etree.fromstring(sp_rect(341 + i * 3, '文本块%d' % (i + 1), bx, by, bw, bh, 'F2F2F2')))
        t.append(etree.fromstring(sp_text(342 + i * 3, bx, by, bw, bh, '文本', 14)))


def draw_p9(t):
    """三张透视斜置卡片 + 平滑弧面地面"""
    # 地面：贝塞尔平滑曲线填充
    ground = ("<p:sp xmlns:p='%s' xmlns:a='%s'><p:nvSpPr><p:cNvPr id='400' name='地面'/>"
              "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x='0' y='5200000'/>"
              "<a:ext cx='%d' cy='%d'/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>"
              "<a:rect l='0' t='0' r='100000' b='100000'/><a:pathLst><a:path w='100000' h='100000'>"
              "<a:moveTo><a:pt x='0' y='45000'/></a:moveTo>"
              "<a:cubicBezTo><a:pt x='25000' y='70000'/><a:pt x='60000' y='88000'/><a:pt x='100000' y='78000'/></a:cubicBezTo>"
              "<a:lnTo><a:pt x='100000' y='100000'/></a:lnTo><a:lnTo><a:pt x='0' y='100000'/></a:lnTo>"
              "<a:close/></a:path></a:pathLst></a:custGeom>"
              "<a:solidFill><a:srgbClr val='ECECEC'/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>"
              "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang='zh-CN'/></a:p></p:txBody></p:sp>"
              % (P, A, 1280 * EMU, 200 * EMU))
    t.append(etree.fromstring(ground))
    # 三张透视卡片：后 → 前，梯形（上窄下宽 + 右侧收窄）
    cards = [
        (1010, 85, 210, 400, 26, 'BCB8B8'),
        (840, 140, 210, 400, 24, 'C6C2C2'),
        (680, 175, 210, 380, 22, 'D0CCCC'),
    ]
    for i, (x, y, w, h, k, col) in enumerate(cards):
        pts = [(x, y + k), (x + w, y), (x + w, y + h - k), (x, y + h)]
        t.append(etree.fromstring(sp_poly(410 + i, '卡片%d' % (i + 1), pts, col)))
        # 厚度边：右侧面 + 顶面（各一条细梯形）
        d = 9
        t.append(etree.fromstring(sp_poly(420 + i, '卡面右%d' % (i + 1), [
            (x + w, y), (x + w + d, y - d // 2), (x + w + d, y + h - k - d // 2), (x + w, y + h - k)], 'A8A4A4')))
        t.append(etree.fromstring(sp_poly(425 + i, '卡面顶%d' % (i + 1), [
            (x, y + k), (x + w, y), (x + w + d, y - d // 2), (x + d, y + k - d // 2)], '9C9898')))
        t.append(etree.fromstring(sp_text(430 + i, x + 20, y + h // 2 - 30, w - 40, 60, '图片', 16)))


DRAWERS = {6: draw_p6, 7: draw_p7, 8: draw_p8, 9: draw_p9, 10: draw_p10}


# ---------- 主流程 ----------
def placeholder(sid, x, y, w, h):
    return ('<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="%d" name="图片占位"/>'
            '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="F7F7F7"/></a:solidFill>'
            '<a:ln w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>'
            '<a:prstDash val="dash"/></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>'
            '<a:r><a:rPr lang="zh-CN" sz="1600"><a:solidFill><a:srgbClr val="8C8C8C"/></a:solidFill>'
            '</a:rPr><a:t>图片</a:t></a:r></a:p></p:txBody></p:sp>'
            % (P, A, sid, x * EMU, y * EMU, w * EMU, h * EMU))


def main():
    z = zipfile.ZipFile(SRC)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    m = {r.get('Id'): r.get('Target') for r in rels}
    parts = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '')
             for s in pres.find('{%s}sldIdLst' % P)]

    stats = {'pic→placeholder': 0, 'drawn': 0}
    for idx, part in enumerate(parts, 1):
        root = etree.fromstring(items[part])
        spTree = root.find('.//{%s}spTree' % P)
        if idx in DRAWERS:
            for ch in list(spTree):
                if etree.QName(ch).localname not in ('nvGrpSpPr', 'grpSpPr'):
                    spTree.remove(ch)
            DRAWERS[idx](spTree)
            stats['drawn'] += 1
        elif idx in PIC_PAGES:
            for pic in list(root.iter('{%s}pic' % P)):
                xf = pic.find('.//{%s}xfrm' % A)
                if xf is None:
                    continue
                off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
                if off is None or ext is None:
                    continue
                x, y = int(off.get('x')) // EMU, int(off.get('y')) // EMU
                w, h = int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU
                if w < 40 or h < 30:
                    continue
                nv = pic.find('.//{%s}cNvPr' % P)
                sid = int(nv.get('id')) if nv is not None and str(nv.get('id', '')).isdigit() else 900
                parent = pic.getparent()
                parent.insert(list(parent).index(pic), etree.fromstring(placeholder(sid, x, y, w, h)))
                parent.remove(pic)
                stats['pic→placeholder'] += 1
            # 大图占位沉底
            for sp in list(spTree):
                nv = sp.find('.//{%s}cNvPr' % P)
                if nv is None or '图片占位' not in (nv.get('name') or ''):
                    continue
                xf = sp.find('.//{%s}xfrm' % A)
                if xf is None:
                    continue
                ext = xf.find('{%s}ext' % A)
                if ext is None:
                    continue
                if (int(ext.get('cx')) // EMU) * (int(ext.get('cy')) // EMU) < AREA * 0.30:
                    continue
                spTree.remove(sp)
                spTree.insert(2 if len(spTree) >= 2 else 0, sp)
        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    print('完整重绘完成：重画 %d 页，图片转占位 %d 个 → %s'
          % (stats['drawn'], stats['pic→placeholder'], OUT))


if __name__ == '__main__':
    main()
