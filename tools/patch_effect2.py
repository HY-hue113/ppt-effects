"""效果2 补丁：底图转幻灯片背景、插入玻璃图文框、校正文字、写入 2 秒平滑切换。

用法:
    python patch_effect2.py --in 导出.pptx --spec ../effects/002_玻璃碎片/spec.json --out 定稿.pptx
"""
import argparse, json, os, zipfile
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'
NS = {'a': A, 'p': P, 'r': R}

FRAME_XML = '''<p:sp xmlns:p="%(P)s" xmlns:a="%(A)s" xmlns:r="%(R)s" useBgFill="1">
 <p:nvSpPr><p:cNvPr id="%(id)d" name="%(name)s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
 <p:spPr>
  <a:xfrm rot="%(rot)d"><a:off x="%(x)d" y="%(y)d"/><a:ext cx="%(cx)d" cy="%(cy)d"/></a:xfrm>
  <a:prstGeom prst="frame"><a:avLst><a:gd name="adj1" fmla="val %(adj)d"/></a:avLst></a:prstGeom>
  <a:ln><a:noFill/></a:ln>
  <a:effectLst><a:outerShdw blurRad="%(blur)d" sx="%(sx)d" sy="%(sy)d" algn="ctr" rotWithShape="0">
   <a:prstClr val="black"><a:alpha val="%(alpha)d"/></a:prstClr></a:outerShdw></a:effectLst>
 </p:spPr>
 <p:style>
  <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>
  <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
  <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
  <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
 </p:style>
</p:sp>'''

BG_XML = ('<p:bg xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:bgPr><a:blipFill dpi="0" rotWithShape="1">'
          '<a:blip r:embed="%s"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
          '<a:effectLst/></p:bgPr></p:bg>')

TRANS = ('<mc:AlternateContent xmlns:mc="%s" xmlns:p="%s" xmlns:p14="%s" xmlns:p159="%s">'
         '<mc:Choice Requires="p159"><p:transition spd="%s" p14:dur="%d"><p159:morph option="%s"/></p:transition></mc:Choice>'
         '<mc:Fallback><p:transition spd="%s"><p:fade/></p:transition></mc:Fallback>'
         '</mc:AlternateContent>')


def sp_by_name(root, name):
    for sp in root.iter('{%s}sp' % P):
        cn = sp.find('.//{%s}cNvPr' % P)
        if cn is not None and cn.get('name') == name:
            return sp
    return None


def set_xfrm(sp, box):
    x, y, cx, cy = box
    xf = sp.find('.//a:xfrm', NS)
    xf.find('a:off', NS).attrib.update({'x': str(x), 'y': str(y)})
    xf.find('a:ext', NS).attrib.update({'cx': str(cx), 'cy': str(cy)})


def style_run(sp, size, font, color):
    for rpr in sp.iter('{%s}rPr' % A):
        rpr.set('sz', str(size))
        for tag in ('latin', 'ea', 'cs'):
            el = rpr.find('a:' + tag, NS)
            if el is None:
                el = etree.SubElement(rpr, '{%s}%s' % (A, tag))
            el.set('typeface', font)
        fill = rpr.find('a:solidFill', NS)
        if fill is None:
            fill = etree.Element('{%s}solidFill' % A)
            rpr.insert(0, fill)
        for c in list(fill):
            fill.remove(c)
        etree.SubElement(fill, '{%s}srgbClr' % A).set('val', color)


def center_box(sp, ln_pct):
    body = sp.find('.//a:bodyPr', NS)
    body.set('anchor', 'ctr')
    for attr in ('lIns', 'rIns', 'tIns', 'bIns'):
        body.set(attr, '0')
    for p in sp.iter('{%s}p' % A):
        ppr = p.find('a:pPr', NS)
        if ppr is None:
            ppr = etree.Element('{%s}pPr' % A)
            p.insert(0, ppr)
        ppr.set('algn', 'ctr')
        for ln in ppr.findall('a:lnSpc', NS):
            ppr.remove(ln)
        ln = etree.Element('{%s}lnSpc' % A)
        etree.SubElement(ln, '{%s}spcPct' % A).set('val', str(ln_pct))
        ppr.insert(0, ln)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--in', dest='src', required=True)
    ap.add_argument('--spec', required=True)
    ap.add_argument('--out', dest='dst', required=True)
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding='utf-8'))
    frames = spec['frames']
    sh = frames['shadow']
    text = spec['text']
    tr = spec['transition']
    ln_pct = 90000

    with zipfile.ZipFile(args.src) as z:
        items = {n: z.read(n) for n in z.namelist()}

    for idx in (1, 2, 3):
        sname = 'ppt/slides/slide%d.xml' % idx
        root = etree.fromstring(items[sname])
        tree = root.find('.//p:spTree', NS)

        anchor = sp_by_name(root, 'Rectangle 3')
        pos = list(tree).index(anchor) + 1
        for i, row in enumerate(frames['pages'][str(idx)]):
            x, y, cx, cy, rot, adj = row
            el = etree.fromstring(FRAME_XML % dict(
                P=P, A=A, R=R, id=100 + i, name=frames['names'][i],
                x=x, y=y, cx=cx, cy=cy, rot=rot, adj=adj,
                blur=sh['blurRad'], sx=sh['sx'], sy=sh['sy'], alpha=sh['alpha']))
            tree.insert(pos + i, el)

        bg_pic = None
        for pic in root.iter('{%s}pic' % P):
            cn = pic.find('.//{%s}cNvPr' % P)
            if cn is not None and cn.get('name') == 'Image 2':
                bg_pic = pic
                break
        if bg_pic is not None:
            embed = bg_pic.find('.//a:blip', NS).get('{%s}embed' % R)
            root.find('.//p:cSld', NS).insert(0, etree.fromstring(BG_XML % (P, A, R, embed)))
            tree.remove(bg_pic)

        if idx == 2:
            tb = sp_by_name(root, 'TextBox 4')
            set_xfrm(tb, text['page2']['box_emu'])
            style_run(tb, text['page2']['size_pt'] * 100, text['font'], text['color'])
            center_box(tb, ln_pct)
            body = tb.find('.//p:txBody', NS)
            paras = body.findall('a:p', NS)
            if len(paras) > 1:
                first = paras[0]
                etree.SubElement(first, '{%s}br' % A)
                for extra in paras[1:]:
                    for child in list(extra):
                        if etree.QName(child).localname in ('r', 'br'):
                            first.append(child)
                    body.remove(extra)

        if idx == 3:
            p3 = text['page3']
            tb = sp_by_name(root, 'TextBox 4')
            set_xfrm(tb, p3['title']['box_emu'])
            style_run(tb, p3['title']['size_pt'] * 100, text['font'], text['color'])
            center_box(tb, ln_pct)
            for i, num in enumerate(p3['numbers']):
                sp = sp_by_name(root, 'TextBox %d' % (5 + i))
                box = list(num['box_emu'])
                box[1] += p3.get('numbers_y_offset', 0)
                set_xfrm(sp, box)
                style_run(sp, num['size_pt'] * 100, text['font'], text['color'])
                center_box(sp, ln_pct)

        for old in [e for e in root if etree.QName(e).localname == 'AlternateContent']:
            root.remove(old)
        if idx in tr['pages']:
            root.append(etree.fromstring(TRANS % (MC, P, P14, P159,
                                                  tr['spd'], tr['p14_dur'], tr['morph_by'], tr['spd'])))

        items[sname] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(args.dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)
    print('patched ->', args.dst)


if __name__ == '__main__':
    main()
