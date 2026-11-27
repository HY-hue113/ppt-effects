# -*- coding: utf-8 -*-
"""把文案与图片填进已处理好的学术模板（调用阶段用）

前提：模板已经是「占位版」（标题 / 文本 / 图片 / 图表 占位），不需要再删页或占位化。

用法：
    python fill_template.py 模板.pptx 内容.json 输出.pptx

内容.json 结构：
{
  "school": "北京大学",                  // 可选
  "pages": {
     "1": {"title": "封面主标题", "texts": ["副标题"], "images": []},
     "6": {"title": "研究背景及现状"},
     "7": {"title": "研究背景",
           "texts": ["国内侧重居家社区协同，2025 年覆盖率 52.3%。", "…"],
           "images": ["img/fig1.png", "img/fig2.png"]}
  }
}
"""
import json
import os
import sys
import zipfile

from lxml import etree

import sys
sys.path.insert(0, r'<LIB_ROOT>\tools')
import typeset  # noqa: E402
from PIL import Image

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships'
IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
EMU = 9525


def set_paras(sp, lines):
    """把一个文本框改写为若干行（沿用原有字体字号）"""
    txBody = sp.find('{%s}txBody' % P)
    if txBody is None:
        return False
    paras = txBody.findall('{%s}p' % A)
    if not paras:
        return False
    proto = paras[0]
    rPr = None
    runs = proto.findall('{%s}r' % A)
    if runs:
        rPr = runs[0].find('{%s}rPr' % A)
    for p in paras[1:]:
        txBody.remove(p)
    for r in proto.findall('{%s}r' % A):
        proto.remove(r)
    for i, line in enumerate(lines):
        para = proto if i == 0 else etree.SubElement(txBody, '{%s}p' % A)
        r = etree.SubElement(para, '{%s}r' % A)
        if rPr is not None:
            r.append(etree.fromstring(etree.tostring(rPr)))
        t = etree.SubElement(r, '{%s}t' % A)
        t.text = line
    return True






def is_nav_shape(sp):
    xf = sp.find('.//{%s}xfrm' % A)
    if xf is None:
        return False
    off, ext = xf.find('{%s}off' % A), xf.find('{%s}ext' % A)
    if off is None or ext is None:
        return False
    x = int(off.get('x')) // 9525
    y = int(off.get('y')) // 9525
    w = int(ext.get('cx')) // 9525
    h = int(ext.get('cy')) // 9525
    if y < 110 and h <= 80 and x > 300:     # 顶部导航（只认右侧横排的标签）
        return True
    if x < 230 and w <= 210 and h <= 90 and y >= 150:    # 左侧导航（排除页标题）
        return True
    return False

def fit_font(sp, text):
    """断行 + 字号二分（下限 20pt）：能放下的最大字号，断行写进 <a:br/>"""
    if not text:
        return
    xf = sp.find('.//{%s}xfrm' % A)
    if xf is None:
        return
    ext = xf.find('{%s}ext' % A)
    if ext is None:
        return
    w = int(ext.get('cx')) // 9525
    h = int(ext.get('cy')) // 9525
    # 原字号作为上限
    top = 32
    for r in sp.iter('{%s}r' % A):
        rPr = r.find('{%s}rPr' % A)
        if rPr is not None and rPr.get('sz'):
            top = max(top, int(rPr.get('sz')) // 100)
    r = typeset.fit_block(text, max(w - 32, 40), max(h - 24, 20),
                          max_pt=min(top, 60), min_pt=20, max_lines=3)
    # 写回：字号统一
    for run in sp.iter('{%s}r' % A):
        rPr = run.find('{%s}rPr' % A)
        if rPr is not None:
            rPr.set('sz', str(int(r['pt'] * 100)))
    # 只改字号：保留原有字体/颜色/加粗（不重建 run，避免丢样式）


def rels_name(part):
    return '%s/_rels/%s.rels' % (os.path.dirname(part), os.path.basename(part))


def fill_image(items, part, sp, img_path, media_idx):
    """把「图片」占位框换成图片填充"""
    name = 'fill%d%s' % (media_idx, os.path.splitext(img_path)[1].lower())
    items['ppt/media/' + name] = open(img_path, 'rb').read()
    ext = os.path.splitext(img_path)[1].lower().lstrip('.')
    ct = items['[Content_Types].xml'].decode('utf-8')
    if ('Extension="%s"' % ext) not in ct:
        ct = ct.replace('</Types>', '<Default Extension="%s" ContentType="image/%s"/></Types>'
                        % (ext, 'jpeg' if ext in ('jpg', 'jpeg') else ext))
        items['[Content_Types].xml'] = ct.encode('utf-8')

    rn = rels_name(part)
    rels = etree.fromstring(items[rn]) if rn in items else etree.fromstring(
        ('<Relationships xmlns="%s"/>' % PKG_R).encode())
    rid = 'rIdFill%d' % media_idx
    rel = etree.SubElement(rels, '{%s}Relationship' % PKG_R)
    rel.set('Id', rid); rel.set('Type', IMG_REL); rel.set('Target', '../media/%s' % name)
    items[rn] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

    spPr = sp.find('{%s}spPr' % P)
    if spPr is None:
        return False
    for old in spPr.findall('{%s}solidFill' % A) + spPr.findall('{%s}blipFill' % A):
        spPr.remove(old)
    for ln in spPr.findall('{%s}ln' % A):
        spPr.remove(ln)
    spPr.append(etree.fromstring(
        '<a:blipFill xmlns:a="%s" xmlns:r="%s"><a:blip r:embed="%s"/>'
        '<a:stretch><a:fillRect/></a:stretch></a:blipFill>' % (A, R, rid)))

    # 清空占位文字
    for tx in sp.iter('{%s}t' % A):
        tx.text = ''
    # 保持原占位框几何不变：图片按框裁切填充（不放大、不偏移）
    return True


def fill(tpl, spec, out):
    data = json.load(open(spec, encoding='utf-8'))
    pages = {int(k): v for k, v in data.get('pages', {}).items()}
    school = data.get('school')
    sections = data.get('sections') or []   # NAV_LEN_HINT
    for _s in sections:
        if len(_s) > 6:
            print('提示：导航名「%s」超过 6 字，可能折行，建议精简' % _s)

    with zipfile.ZipFile(tpl) as z:
        items = {n: z.read(n) for n in z.namelist()}

    pres = etree.fromstring(items['ppt/presentation.xml'])
    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    m = {r.get('Id'): r.get('Target') for r in rels}
    order = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '') for s in pres.find('{%s}sldIdLst' % P)]

    media_idx = 0
    for idx, part in enumerate(order, 1):
        if part not in items:
            continue
        conf = pages.get(idx)
        if conf is None:
            continue
        root = etree.fromstring(items[part])
        texts = list(conf.get('texts', []))
        best_title_area = 0
        images = list(conf.get('images', []))
        txt_i = 0

        # 校名
        if school:
            for sp in root.iter('{%s}sp' % P):
                t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
                if t == '学校名称':
                    set_paras(sp, [school])

        is_nav = is_nav_shape

        # 导航栏章节名（若提供了 sections）
        if sections:
            nav_items = []
            for sp in root.iter('{%s}sp' % P):
                if not is_nav(sp):
                    continue
                t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
                if t not in ('标题', '文本'):
                    continue
                xf = sp.find('.//{%s}xfrm' % A)
                off = xf.find('{%s}off' % A) if xf is not None else None
                if off is None:
                    continue
                nav_items.append((int(off.get('x')), int(off.get('y')), sp))
            nav_items.sort(key=lambda v: (v[1], v[0]))   # 顶部按 x、左侧按 y
            # 去掉被校徽遮住的装饰项（左栏顶部 y<250）
            nav_items = [n for n in nav_items if not (n[0] < 260 and n[1] < 250)]
            off_n = int(data.get('nav_offset') or 0)      # 从第几个导航框开始填
            if off_n:
                nav_items = nav_items[off_n:]
            if len(nav_items) > len(sections):
                nav_items = nav_items[:len(sections)]
            for i, (_, _, sp) in enumerate(nav_items):
                if i >= len(sections):
                    continue
                label = sections[i]
                set_paras(sp, [label])
                # 按宽度缩小字号，避免溢出换行
                xf2 = sp.find('.//{%s}xfrm' % A)
                ext2 = xf2.find('{%s}ext' % A) if xf2 is not None else None
                if ext2 is None:
                    continue
                avail = int(ext2.get('cx')) // 9525 - 12
                for r in sp.iter('{%s}r' % A):
                    rPr = r.find('{%s}rPr' % A)
                    if rPr is None:
                        continue
                    cur = int(rPr.get('sz') or 1800)
                    fit = int(avail / max(1, len(label)) * 100)
                    if fit < cur:
                        rPr.set('sz', str(max(800, fit)))

        def _pos(sp):
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                return (9999, 9999)
            off = xf.find('{%s}off' % A)
            if off is None:
                return (9999, 9999)
            return (int(off.get('y')) // 9525, int(off.get('x')) // 9525)

        ordered = sorted(list(root.iter('{%s}sp' % P)), key=_pos)
        for sp in ordered:
            t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
            if is_nav(sp):          # 导航栏标签不参与标题/正文填充
                continue
            if t == '标题' and conf.get('title'):
                # 每页只填面积最大的那个标题（页面主标题），其余保留占位
                xf0 = sp.find('.//{%s}xfrm' % A)
                area = 0
                if xf0 is not None:
                    e0 = xf0.find('{%s}ext' % A)
                    if e0 is not None:
                        area = int(e0.get('cx')) * int(e0.get('cy'))
                if area >= best_title_area and area > 0:
                    best_title_area = area
                    set_paras(sp, [conf['title']])
            elif t == '文本':
                if txt_i < len(texts):
                    set_paras(sp, [texts[txt_i]])
                    txt_i += 1
            elif t == '图片' and images:
                img = images.pop(0)
                media_idx += 1
                fill_image(items, part, sp, img, media_idx)

        # TITLE_SORTED：面积最大=主标题，次大=副标题
        titles = []
        for sp in root.iter('{%s}sp' % P):
            if is_nav(sp):
                continue
            txt = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
            if not txt:
                continue
            xf0 = sp.find('.//{%s}xfrm' % A)
            if xf0 is None:
                continue
            e0 = xf0.find('{%s}ext' % A)
            if e0 is None:
                continue
            titles.append((int(e0.get('cx')) * int(e0.get('cy')), sp, txt))
        # 页面主标题 = 位置最靠上（y 最小）的那个标题形状；灰条等次级标题留空
        def _ty(sp):
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                return 9999
            off = xf.find('{%s}off' % A)
            return int(off.get('y')) // 9525 if off is not None else 9999
        def _tw(sp):
            xf = sp.find('.//{%s}xfrm' % A)
            if xf is None:
                return 0
            ext = xf.find('{%s}ext' % A)
            return int(ext.get('cx')) // 9525 if ext is not None else 0
        # 内容页：窄框（页标题）当主标题；宽条（灰底副标题条）留空
        narrow = [t for t in titles if _tw(t[1]) < 500]
        wide = [t for t in titles if _tw(t[1]) >= 500]
        titles = narrow + wide
        titles.sort(key=lambda v: (_tw(v[1]) >= 500, _ty(v[1])))
        if conf.get('title') and titles:
            main = titles[0][1]
            set_paras(main, [conf['title']])
            fit_font(main, conf['title'])
            if len(titles) > 1 and conf.get('subtitle'):
                sub = max(titles[1:], key=lambda v: v[0])[1]
                set_paras(sub, [conf['subtitle']])
                fit_font(sub, conf['subtitle'])
            elif len(titles) > 1:
                # C：内容页次级标题留空
                tx = max(titles[1:], key=lambda v: v[0])[1].find('{%s}txBody' % P)
                if tx is not None:
                    for para in tx.findall('{%s}p' % A):
                        for r in para.findall('{%s}r' % A):
                            for t in r.findall('{%s}t' % A):
                                t.text = '' 

        items[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return len(pages)


if __name__ == '__main__':
    n = fill(sys.argv[1], sys.argv[2], sys.argv[3])
    print('已填充 %d 页 → %s' % (n, os.path.basename(sys.argv[3])))
