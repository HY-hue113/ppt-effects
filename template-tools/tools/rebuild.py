# -*- coding: utf-8 -*-
"""学术模板重建（从头来过，按用户指定的 6 步）

1. 删除第 1–3 页
2. 其余页面原样保留
3. 封面/结尾页：只改标题；内容页：导航栏→「标题」，页面标题→「标题」，其余文字→「文本」
4. 图片素材删除，原位留「图片」占位框
5. 图表素材删除，原位留「图表」占位框
6. 校徽/校名：母版+版式里的 logo 换成「校徽」占位（对应 PDF 的母版替换法）

用法：python rebuild.py <源.pptx> <输出.pptx>
"""
import copy
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
EMU = 9525
DROP_PAGES = 3


def q(tag):
    return '{%s}%s' % (P, tag)


def get_box(el):
    xf = el.find('.//{%s}xfrm' % A) if el.find('.//{%s}xfrm' % A) is not None else el.find('.//{%s}xfrm' % P)
    if xf is None:
        return None
    off = xf.find('{%s}off' % A) if xf.find('{%s}off' % A) is not None else xf.find('{%s}off' % P)
    ext = xf.find('{%s}ext' % A) if xf.find('{%s}ext' % A) is not None else xf.find('{%s}ext' % P)
    if off is None or ext is None:
        return None
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def abs_box(el):
    """含组合变换的绝对坐标（px）"""
    xf = el.find('.//{%s}xfrm' % A) if el.find('.//{%s}xfrm' % A) is not None else el.find('.//{%s}xfrm' % P)
    if xf is None:
        return None
    off = xf.find('{%s}off' % A) if xf.find('{%s}off' % A) is not None else xf.find('{%s}off' % P)
    ext = xf.find('{%s}ext' % A) if xf.find('{%s}ext' % A) is not None else xf.find('{%s}ext' % P)
    if off is None or ext is None:
        return None
    x, y = int(off.get('x')), int(off.get('y'))
    w, h = int(ext.get('cx')), int(ext.get('cy'))
    cur = el.getparent()
    while cur is not None:
        if etree.QName(cur).localname == 'grpSp':
            gpr = cur.find(q('grpSpPr'))
            gx = gpr.find('{%s}xfrm' % A) if gpr is not None else None
            if gx is not None:
                go, ge = gx.find('{%s}off' % A), gx.find('{%s}ext' % A)
                co, ce = gx.find('{%s}chOff' % A), gx.find('{%s}chExt' % A)
                if None not in (go, ge, co, ce):
                    gx0, gy0 = int(go.get('x')), int(go.get('y'))
                    gw, gh = int(ge.get('cx')), int(ge.get('cy'))
                    cx0, cy0 = int(co.get('x')), int(co.get('y'))
                    cw, ch = int(ce.get('cx')), int(ce.get('cy'))
                    if cw and ch:
                        sx, sy = gw / cw, gh / ch
                        x = gx0 + (x - cx0) * sx
                        y = gy0 + (y - cy0) * sy
                        w, h = w * sx, h * sy
        cur = cur.getparent()
    return (int(x) // EMU, int(y) // EMU, int(w) // EMU, int(h) // EMU)


def set_text(sp, text):
    """逐段改写，保留字体/字号与段落数"""
    txBody = sp.find(q('txBody'))
    if txBody is None:
        return False
    paras = txBody.findall('{%s}p' % A)
    if not paras:
        return False
    for para in paras:
        runs = para.findall('{%s}r' % A)
        rPr = runs[0].find('{%s}rPr' % A) if runs else None
        for r in runs:
            para.remove(r)
        for br in para.findall('{%s}br' % A):
            para.remove(br)
        r = etree.SubElement(para, '{%s}r' % A)
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        t = etree.SubElement(r, '{%s}t' % A)
        t.text = text
    return True


def placeholder(box, label, sid):
    x, y, w, h = box
    xml = (
        '<p:sp xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
        '<p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="F2F5F9"/></a:solidFill>'
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="AFC0D4"/></a:solidFill>'
        '<a:prstDash val="dash"/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="1600">'
        '<a:solidFill><a:srgbClr val="7C8B9C"/></a:solidFill></a:rPr>'
        '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (P, A, R, sid, label + '占位', x * EMU, y * EMU, w * EMU, h * EMU, label))
    return etree.fromstring(xml)


def max_sz(sp):
    szs = [int(v) for v in re.findall(rb'sz="(\d+)"', etree.tostring(sp))]
    return max(szs) if szs else 0


def nav_band(box):
    if not box:
        return False
    x, y, w, h = box
    return (y < 120) or (x < 170 and w <= 170 and h <= 220)


def scrub_page(root, kind, sid_base):
    """kind: cover / end / content"""
    spTree = root.find('.//{%s}spTree' % P)
    if spTree is None:
        return
    removed_imgs, removed_charts, texts = 0, 0, 0

    # 4) 图片 → 图片占位（背景大图保留）
    for pic in list(root.iter(q('pic'))):
        b = abs_box(pic)
        if not b:
            continue
        if b[2] >= 1100:
            continue
        parent = pic.getparent()
        parent.remove(pic)
        spTree.append(placeholder(b, '图片', sid_base + 10 + removed_imgs))
        removed_imgs += 1

    # 5) 图表 → 图表占位
    for gf in list(root.iter(q('graphicFrame'))):
        b = abs_box(gf)
        if not b:
            continue
        is_chart = gf.find('.//{%s}chart' % C) is not None
        if not is_chart:
            continue
        parent = gf.getparent()
        parent.remove(gf)
        spTree.append(placeholder(b, '图表', sid_base + 100 + removed_charts))
        removed_charts += 1

    # 3) 文字
    page_max = max([max_sz(sp) for sp in root.iter(q('sp'))] or [0])
    for sp in list(root.iter(q('sp'))):
        if sp.find(q('txBody')) is None:
            continue
        txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
        if not txt:
            continue
        if txt in ('图片', '图表', '校徽'):      # 占位框不参与文字替换
            continue
        sz = max_sz(sp)
        b = get_box(sp)

        if kind in ('cover', 'end'):
            if sz >= max(2400, page_max) and sz >= page_max * 0.95:
                set_text(sp, '标题')
                texts += 1
            continue                      # 封面/结尾：其余文字原样

        # 内容页
        if nav_band(b):
            set_text(sp, '标题')
        elif sz >= 2400 and sz >= page_max * 0.9:
            set_text(sp, '标题')
        else:
            set_text(sp, '文本')
        texts += 1

    return removed_imgs, removed_charts, texts


def delete_first_pages(items, n):
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    lst = pres.find(q('sldIdLst'))
    ids = lst.findall(q('sldId'))
    drop_rids, drop_parts = [], []
    for el in ids[:n]:
        rid = el.get('{%s}id' % R)
        drop_rids.append(rid)
        for rel in rels:
            if rel.get('Id') == rid:
                tgt = rel.get('Target').replace('../', '')
                drop_parts.append('ppt/' + tgt)
                rels.remove(rel)
        lst.remove(el)
    items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True, encoding='UTF-8', standalone=True)
    items['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

    for part in drop_parts:
        items.pop(part, None)
        items.pop(part.replace('slides/', 'slides/_rels/') + '.rels', None)
        # Content_Types 里去掉 override
        ct = items['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override PartName="/%s"[^/]*/>' % re.escape(part), '', ct)
        items['[Content_Types].xml'] = ct.encode('utf-8')
    return drop_parts


def order_parts(items):
    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    m = {r.get('Id'): r.get('Target') for r in rels}
    return ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '') for s in pres.find(q('sldIdLst'))]


def rebuild(src, dst):
    shutil.copyfile(src, dst)
    z = zipfile.ZipFile(dst)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()

    delete_first_pages(items, DROP_PAGES)          # 1) 删前三页
    pages = order_parts(items)

    for i, part in enumerate(pages):               # 2)+3) 逐页处理
        if part not in items:
            continue
        root = etree.fromstring(items[part])
        if i < 5:
            kind = 'cover'                      # 前五页：只改标题
        elif i == len(pages) - 1:
            kind = 'end'
        else:
            kind = 'content' 
        scrub_page(root, kind, sid_base=3000 + i * 200)
        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 6) 母版/版式 logo → 「校徽」占位（PDF 方法：一次全换）
    for n in [x for x in items if re.match(r'ppt/(slideMasters/slideMaster\d+|slideLayouts/slideLayout\d+)\.xml$', x)]:
        root = etree.fromstring(items[n])
        spTree = root.find('.//{%s}spTree' % P)
        pics = list(root.iter(q('pic')))
        for k, pic in enumerate(pics):
            b = get_box(pic)
            if not b or b[2] >= 1100:
                continue
            parent = pic.getparent()
            parent.remove(pic)
            if spTree is not None:
                spTree.append(placeholder(b, '校徽', 9000 + k))
        items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return len(pages)


if __name__ == '__main__':
    n = rebuild(sys.argv[1], sys.argv[2])
    print('完成：保留 %d 页' % n)
