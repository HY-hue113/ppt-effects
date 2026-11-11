"""效果11 · a 张图（严格对应 + 复制页）

- 页数 = a + 1：p1 封面；p2..p(a) 每页带「图片(n-1)+图片n」；p(a+1) 末尾单图
- 页面骨架：p2（图片在右）与 p3（图片在左）交替复制
- 严格对应：小四边形副本K图 ← 第 K 张图；小四边形副本K = 色块 K
- 背景色 = 该页层级最高的纯色色块颜色
"""
import os, re, copy, zipfile, shutil
from lxml import etree

SRC = r'<USER_HOME>\<OUTPUT_DIR>\效果11-模板套用.pptx'
OUT = r'<USER_HOME>\<OUTPUT_DIR>\效果11-多图.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

IMGS = [
    r'<DSH_ATTACHMENTS>\v1\objects\be\beeab8ad8c130c6da6560e4626f40ffd7ffdb2dd4211bc322a68ee03e1238f31',
    r'<DSH_ATTACHMENTS>\v1\objects\79\79c7ba221f9a49bff5b4a91e7ea2b3af9cb7241aa2dc1459baebfd6e128d0f6c',
    r'<DSH_ATTACHMENTS>\v1\objects\53\539c7a3d594ca9f02d5c1c6126358d41ff835ed6c8af0801e8b1f2ceb92aeda7',
    r'<DSH_ATTACHMENTS>\v1\objects\3c\3c18d102d118f36798c7c1808e20ab299ea72e283a4c7fb9ed9ebf8f74ccff65',
    r'<DSH_ATTACHMENTS>\v1\objects\a8\a8eb323b40a77068b37401a8a34cfff3a6e9aff7542b01ef20f4a536e441ffd9',
    r'<DSH_ATTACHMENTS>\v1\objects\46\469e386a512252f3920a9fc9d1fbbaafcf0760a8d520952a700341f8926b2633',
    r'<DSH_ATTACHMENTS>\v1\objects\e9\e9466e1a860cc7cf36247873fad1fecaf32eed76615d522ac3ce24d95e2e5a8b',
    r'<DSH_ATTACHMENTS>\v1\objects\fd\fde252a7f2fe9f118933c5ce7c3b38be969b6cd37634ef14ed21615628998975',
    r'<DSH_ATTACHMENTS>\v1\objects\0d\0d3944e073c337e85bdf488e2650367d661045bc068c10524fbc8bcc485a7bf6',
]
COLORS = ['3A1113', 'C35D09', '1F3A5F', '2E5A3A', '6B2F5B', '8A5A00', '1E6B6B', '7A2E2E', '4B4B8A', '2F6B3A']
A_N = len(IMGS)

z = zipfile.ZipFile(SRC)
items = {n: z.read(n) for n in z.namelist()}
z.close()

pres = etree.fromstring(items['ppt/presentation.xml'])
pres_rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
types = etree.fromstring(items['[Content_Types].xml'])
sld_lst = pres.find('{%s}sldIdLst' % P)

# 现有 slide -> rId 映射
rel_by_target = {r.get('Target'): r.get('Id') for r in pres_rels if r.get('Target', '').startswith('slides/')}

def slide_xml(n):
    return items['ppt/slides/slide%d.xml' % n]

def slide_rels(n):
    return items['ppt/slides/_rels/slide%d.xml.rels' % n]

next_slide = max(int(m.group(1)) for m in (re.match(r'ppt/slides/slide(\d+)\.xml$', k) for k in items) if m) + 1
next_rid_num = max(int(r.get('Id')[3:]) for r in pres_rels) + 1
next_sid = max(int(e.get('id')) for e in sld_lst) + 1

def add_page_from(src_n):
    """复制源页，返回新页号"""
    global next_slide, next_rid_num, next_sid
    n = next_slide
    next_slide += 1
    items['ppt/slides/slide%d.xml' % n] = slide_xml(src_n)
    items['ppt/slides/_rels/slide%d.xml.rels' % n] = slide_rels(src_n)
    rid = 'rId%d' % next_rid_num
    next_rid_num += 1
    e = etree.SubElement(pres_rels, '{%s}Relationship' % PR)
    e.set('Id', rid)
    e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
    e.set('Target', 'slides/slide%d.xml' % n)
    sid = etree.SubElement(sld_lst, '{%s}sldId' % P)
    sid.set('id', str(next_sid))
    sid.set('{%s}id' % R, rid)
    next_sid += 1
    ov = etree.SubElement(types, '{%s}Override' % CT)
    ov.set('PartName', '/ppt/slides/slide%d.xml' % n)
    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
    return n

# 页序：p1(1) + p2..p(a) + 末尾页
plan = [(1, 1, 1)]
for k in range(2, A_N + 1):
    src = 2 if k % 2 == 0 else 3
    plan.append((add_page_from(src), src, k))
plan.append((add_page_from(5), 5, A_N))

# 逐页填充
for idx, (page_n, src_n, k) in enumerate(plan, start=1):
    sname = 'ppt/slides/slide%d.xml' % page_n
    rels_name = 'ppt/slides/_rels/slide%d.xml.rels' % page_n
    root = etree.fromstring(items[sname])
    rels = etree.fromstring(items[rels_name])
    next_rid = max(int(r.get('Id')[3:]) for r in rels) + 1
    # 把从源页复制来的形状改名为目标索引
    if src_n == 5:
        old_cur, old_prev = 4, 3
    else:
        old_cur, old_prev = src_n, max(1, src_n - 1)
    for cNv in root.iter('{%s}cNvPr' % P):
        nm = cNv.get('name', '')
        m = re.fullmatch(r'小四边形副本(\d+)(图?)', nm)
        if not m:
            continue
        num = int(m.group(1)); suf = m.group(2)
        if num == old_cur:
            cNv.set('name', '小四边形副本%d%s' % (k, suf))
        elif num == old_prev:
            cNv.set('name', '小四边形副本%d%s' % (max(1, k - 1), suf))

    for sp in root.iter('{%s}sp' % P):
        cNv = sp.find('.//{%s}cNvPr' % P)
        if cNv is None:
            continue
        nm = cNv.get('name', '')
        spPr = sp.find('{%s}spPr' % P)
        if spPr is None:
            continue

        m_pic = re.fullmatch(r'小四边形副本(\d+)图', nm)
        m_col = re.fullmatch(r'小四边形副本(\d+)', nm)
        if m_pic:
            k = int(m_pic.group(1))
            if 1 <= k <= A_N:
                blip = spPr.find('{%s}blipFill/{%s}blip' % (A, A))
                if blip is not None:
                    media = 'ppt/media/eff11m_%d.jpg' % k
                    if media not in items:
                        items[media] = open(IMGS[k - 1], 'rb').read()
                    rid = 'rId%d' % next_rid
                    next_rid += 1
                    e = etree.SubElement(rels, '{%s}Relationship' % PR)
                    e.set('Id', rid)
                    e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
                    e.set('Target', '../media/' + os.path.basename(media))
                    blip.set('{%s}embed' % R, rid)
        elif m_col:
            # 彻底清空填充（含 schemeClr/gradFill/blipFill），再插入纯色
            for e in list(spPr):
                if etree.QName(e).localname in ('solidFill','gradFill','blipFill','noFill','pattFill'):
                    spPr.remove(e)
            sf = etree.Element('{%s}solidFill' % A)
            c = etree.SubElement(sf, '{%s}srgbClr' % A)
            c.set('val', COLORS[(idx - 1) % len(COLORS)])
            # 插到 prstGeom 之后
            anchor = None
            for e in spPr:
                if etree.QName(e).localname in ('prstGeom','custGeom'):
                    anchor = e
            if anchor is not None:
                anchor.addnext(sf)
            else:
                spPr.append(sf)

# 文字编号跟随页序
    num = max(1, k - 1) if idx < len(plan) else A_N
    for t in root.iter('{%s}t' % A):
        if t.text and re.search(r'\d', t.text):
            t.text = re.sub(r'\d+', str(num), t.text)

    # 末页：纯色块在左、图片块在右（对齐 p2 的层级）
    if idx == len(plan) and A_N % 2 == 1:
        pic_sp = col_sp = None
        for sp in root.iter('{%s}sp' % P):
            nm = sp.find('.//{%s}cNvPr' % P).get('name', '')
            spPr2 = sp.find('{%s}spPr' % P)
            if spPr2 is None:
                continue
            xf = spPr2.find('{%s}xfrm/{%s}off' % (A, A))
            if xf is None:
                continue
            if re.fullmatch(r'小四边形副本\d+图', nm):
                pic_sp = xf
            elif re.fullmatch(r'小四边形副本\d+', nm):
                col_sp = xf
        if pic_sp is not None and col_sp is not None:
            a_x = pic_sp.get('x')
            b_x = col_sp.get('x')
            pic_sp.set('x', b_x)
            col_sp.set('x', a_x)
        # 末页：右侧文字组（grpSp）也移到左边
        for g in root.iter('{%s}grpSp' % P):
            cNv2 = g.find('.//{%s}cNvPr' % P)
            if cNv2 is None:
                continue
            nm2 = cNv2.get('name', '')
            if '组合' not in nm2:
                continue
            gPr = g.find('{%s}grpSpPr' % P)
            xf4 = gPr.find('{%s}xfrm/{%s}off' % (A, A)) if gPr is not None else None
            if xf4 is not None:
                xf4.set('x', '500000')

    # 背景 = 本页主色
    sx = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True).decode('utf-8')
    page_color = COLORS[(idx - 1) % len(COLORS)]
    if '<p:bg>' in sx:
        sx = re.sub(r'(<p:bg>.*?<a:srgbClr val=")[0-9A-Fa-f]{6}(")',
                    lambda m: m.group(1) + page_color + m.group(2), sx, count=1, flags=re.S)
    items[sname] = sx.encode('utf-8')
    items[rels_name] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

# sldIdLst 顺序 = plan
order = []
for e in sld_lst:
    order.append(e)
want = []
for page_n, _src, _k in plan:
    for e in order:
        rid = e.get('{%s}id' % R)
        tgt = next((r.get('Target') for r in pres_rels if r.get('Id') == rid), None)
        if tgt == 'slides/slide%d.xml' % page_n:
            want.append(e)
            break
for e in list(sld_lst):
    sld_lst.remove(e)
for e in want:
    sld_lst.append(e)

# —— 最终编号修正（覆盖复制页继承的旧编号）——
for idx2, (page_n2, src2, k2) in enumerate(plan, start=1):
    sname2 = 'ppt/slides/slide%d.xml' % page_n2
    root2 = etree.fromstring(items[sname2])
    num2 = max(1, idx2 - 1)
    for t2 in root2.iter('{%s}t' % A):
        if t2.text and re.search(r'\d', t2.text):
            t2.text = re.sub(r'\d+', str(num2), t2.text)
    items[sname2] = etree.tostring(root2, xml_declaration=True, encoding='UTF-8', standalone=True)

items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True, encoding='UTF-8', standalone=True)
items['ppt/_rels/presentation.xml.rels'] = etree.tostring(pres_rels, xml_declaration=True, encoding='UTF-8', standalone=True)
items['[Content_Types].xml'] = etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for n, b in items.items():
        zo.writestr(n, b)
print('页数', len(plan), '| done', OUT, round(os.path.getsize(OUT) / 1024), 'KB')
