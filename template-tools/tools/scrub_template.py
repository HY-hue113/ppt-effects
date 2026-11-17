# -*- coding: utf-8 -*-
"""学术模板复刻：删前 3 页 → 文字占位化 → 图片/表格留位 → 校徽/校名可替换

用法：
    python scrub_template.py <源.pptx> <输出.pptx>
"""
import copy
import json
import os
import re
import shutil
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525

KEEP_LABELS = (
    "指导教师", "指导老师", "导师", "答辩人", "答辩日期", "日期", "课题小组",
    "研究方向", "专业", "姓名", "学号", "院系", "学院：", "班级", "电话",
    "邮箱", "成员", "汇报人", "作者", "单位", "职称",
)

TITLE_MIN_SZ = 2600      # ≥26pt 视为标题
LOGO_MAX = 220           # 小于 220×220 且在上部 → 校徽


def set_text(sp, text):
    """逐段改写为 text（保留每段原有字体/字号/颜色与段落数量）"""
    txBody = sp.find('{%s}txBody' % P)
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


def max_font_sz(sp):
    szs = [int(v) for v in re.findall(rb'sz="(\d+)"', etree.tostring(sp))]
    return max(szs) if szs else 0




def abs_xfrm(el):
    """把（可能位于组内的）形状坐标换算成页面坐标。

    返回 (x, y, w, h) 单位 px；遇到无法解析时返回 None。
    """
    xfrm = el.find('.//{%s}xfrm' % A)
    if xfrm is None:
        return None
    off = xfrm.find('{%s}off' % A)
    ext = xfrm.find('{%s}ext' % A)
    if off is None or ext is None:
        return None
    x, y = int(off.get('x')), int(off.get('y'))
    w, h = int(ext.get('cx')), int(ext.get('cy'))

    cur = el.getparent()
    while cur is not None:
        if etree.QName(cur).localname == 'grpSp':
            gpr = cur.find('{%s}grpSpPr' % P)
            gx = gpr.find('.//{%s}xfrm' % A) if gpr is not None else None
            if gx is not None:
                goff = gx.find('{%s}off' % A)
                gext = gx.find('{%s}ext' % A)
                choff = gx.find('{%s}chOff' % A)
                chext = gx.find('{%s}chExt' % A)
                if None not in (goff, gext, choff, chext):
                    gx0, gy0 = int(goff.get('x')), int(goff.get('y'))
                    gw, gh = int(gext.get('cx')), int(gext.get('cy'))
                    cx0, cy0 = int(choff.get('x')), int(choff.get('y'))
                    cw, ch = int(chext.get('cx')), int(chext.get('cy'))
                    if cw and ch:
                        sx, sy = gw / cw, gh / ch
                        x = gx0 + (x - cx0) * sx
                        y = gy0 + (y - cy0) * sy
                        w, h = w * sx, h * sy
        cur = cur.getparent()
    return x // EMU, y // EMU, w // EMU, h // EMU




def find_template_box(root, box, tol=24):
    """在页面上找模板自带的虚线占位框（文字为 正文/图片/图表），按位置匹配"""
    bx, by, bw, bh = box
    best = None
    for sp in root.iter('{%s}sp' % P):
        txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
        if txt not in ('正文', '图片', '图表', '表格'):
            continue
        xfrm = sp.find('.//{%s}xfrm' % A)
        if xfrm is None:
            continue
        off = xfrm.find('{%s}off' % A)
        ext = xfrm.find('{%s}ext' % A)
        if off is None or ext is None:
            continue
        x, y = int(off.get('x')) // EMU, int(off.get('y')) // EMU
        w, h = int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU
        if abs(x - bx) <= tol and abs(y - by) <= tol:
            score = abs(x - bx) + abs(y - by) + abs(w - bw) + abs(h - bh)
            if best is None or score < best[0]:
                best = (score, sp)
    return best[1] if best else None

def make_placeholder(sp, text, box=None):
    """把图片/表格换成圆角矩形 + 占位文字（同位置同尺寸）"""
    if box is not None:
        x = int(box[0] * EMU); y = int(box[1] * EMU)
        cx = int(box[2] * EMU); cy = int(box[3] * EMU)
    else:
        xfrm = sp.find('.//{%s}xfrm' % A)
        if xfrm is None:
            xfrm = sp.find('.//{%s}xfrm' % P)      # graphicFrame 用 p:xfrm
        if xfrm is None:
            return None
        off = xfrm.find('{%s}off' % A)
        if off is None:
            off = xfrm.find('{%s}off' % P)
        ext = xfrm.find('{%s}ext' % A)
        if ext is None:
            ext = xfrm.find('{%s}ext' % P)
        if off is None or ext is None:
            return None
        x, y = off.get('x'), off.get('y')
        cx, cy = ext.get('cx'), ext.get('cy')
    nv = sp.find('.//{%s}cNvPr' % P) if sp is not None else None
    sid = nv.get('id') if nv is not None else '900'
    xml = (
        '<p:sp xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
        '<p:nvSpPr><p:cNvPr id="%s" name="%s占位"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="F2F5F9"/></a:solidFill>'
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="AFC0D4"/></a:solidFill>'
        '<a:prstDash val="dash"/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="zh-CN" sz="1600">'
        '<a:solidFill><a:srgbClr val="7C8B9C"/></a:solidFill></a:rPr>'
        '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
        % (P, A, R, sid, text, x, y, cx, cy, text))
    return etree.fromstring(xml)


KEEP_PAGE_KEYWORDS = ('目录',)




def relabel_placeholders(root, pic_boxes, chart_boxes):
    """把与图片/图表位置重合的「正文」框改写为「图片」/「图表」"""
    n = 0
    for label, boxes in (("图片", [b for b in pic_boxes if b]), ("图表", [b for b in chart_boxes if b])):
        for box in boxes:
            bx, by, bw, bh = box
            for sp in list(root.iter('{%s}sp' % P)):
                txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
                if txt not in ('正文', '图片', '图表'):
                    continue
                xfrm = sp.find('.//{%s}xfrm' % A)
                if xfrm is None:
                    continue
                off = xfrm.find('{%s}off' % A)
                ext = xfrm.find('{%s}ext' % A)
                if off is None or ext is None:
                    continue
                x, y = int(off.get('x')) // EMU, int(off.get('y')) // EMU
                w, h = int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU
                if abs(x - bx) <= 40 and abs(y - by) <= 40 and abs(w - bw) <= 80 and abs(h - bh) <= 80:
                    if set_text(sp, label):
                        n += 1
                    break
    return n

def scrub_slide(root):
    stats = dict(title=0, body=0, pic=0, logo=0, table=0)
    changed = []

    # 图片 / 表格 → 占位
    def in_group(el):
        cur = el.getparent()
        while cur is not None:
            if etree.QName(cur).localname == 'grpSp':
                return True
            cur = cur.getparent()
        return False

    spTree = root.find('.//{%s}spTree' % P)
    _pic_boxes, _chart_boxes = [], []

    # 由多张小图拼成的"组合图"整体换成「图片」占位
    for g in list(root.iter('{%s}grpSp' % P)):
        if not g.findall('.//{%s}pic' % P):
            continue
        b = abs_xfrm(g)
        if not b:
            continue
        w, h = b[2], b[3]
        if 200 <= w < 1100 and h >= 150:
            ph = make_placeholder(g, "图片", box=b)
            if ph is not None:
                parent = g.getparent()
                parent.remove(g)
                (spTree if spTree is not None else parent).append(ph)
                _pic_boxes.append(b)
                stats['pic'] += 1

    for pic in list(root.iter('{%s}pic' % P)):
        box = abs_xfrm(pic)
        if box is None:
            continue
        x, y, w, h = box
        if w >= 1100:              # 铺满整幅的背景/装饰图，保留原样
            continue
        if w < 200 or h < 150:      # 小图标/装饰，保留原样（不占位）
            continue
        _pic_boxes.append(box)
        top_zone = y < 210
        if w < 120 and h < 120 and not top_zone:   # 页面中部的小装饰图标保留
            continue
        is_logo = top_zone
        label = "校徽" if is_logo else "图片"
        tpl = find_template_box(root, box)
        if tpl is not None and not is_logo:
            set_text(tpl, label)
            pic.getparent().remove(pic)
            stats['pic'] += 1
            continue
        ph = make_placeholder(pic, label, box=box)
        if ph is None:
            continue
        parent = pic.getparent()
        idx = list(parent).index(pic)
        parent.remove(pic)
        parent.insert(min(idx, len(list(parent))), ph)
        stats['logo' if is_logo else 'pic'] += 1

    for gf in list(root.iter('{%s}graphicFrame' % P)):
        is_chart = gf.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart') is not None \
            or 'chart' in etree.tostring(gf).decode('utf-8', 'ignore')
        label = "图表" if is_chart else "表格"
        xf2 = gf.find('.//{%s}xfrm' % A)
        if xf2 is None:
            xf2 = gf.find('.//{%s}xfrm' % P)
        tpl = None
        if xf2 is not None:
            o2 = xf2.find('{%s}off' % A) if xf2.find('{%s}off' % A) is not None else xf2.find('{%s}off' % P)
            e2 = xf2.find('{%s}ext' % A) if xf2.find('{%s}ext' % A) is not None else xf2.find('{%s}ext' % P)
            if o2 is not None and e2 is not None:
                _cb = (int(o2.get('x')) // EMU, int(o2.get('y')) // EMU,
                       int(e2.get('cx')) // EMU, int(e2.get('cy')) // EMU)
                _chart_boxes.append(_cb)
                tpl = find_template_box(root, _cb)
        if tpl is not None:
            set_text(tpl, label)
            gf.getparent().remove(gf)
            stats['table'] += 1
            continue
        ph = make_placeholder(gf, label)
        if ph is None:
            continue
        parent = gf.getparent()
        idx = list(parent).index(gf)
        parent.remove(gf)
        parent.insert(idx, ph)
        stats['table'] += 1

    # 文字 → 标题 / 小标题 / 正文
    is_toc = any((''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip() == '目录')
                 for sp in root.iter('{%s}sp' % P))
    for sp in list(root.iter('{%s}sp' % P)):
        if sp.find('{%s}txBody' % P) is None:
            continue
        txt = ''.join(t.text or '' for t in sp.iter('{%s}t' % A)).strip()
        if not txt:
            continue
        if txt == '目录':          # 目录页大标题保留
            continue
        if is_toc:
            xf0 = sp.find('.//{%s}xfrm' % A)
            w0 = int(xf0.find('{%s}ext' % A).get('cx')) // EMU if xf0 is not None and xf0.find('{%s}ext' % A) is not None else 0
            if w0 >= 200:
                set_text(sp, "标题")
                stats['title'] += 1
            continue
        # 仅"纯日期/时间表"保留原文（含年份的叙述句要占位）
        if re.fullmatch(r'\s*\d{4}\s*[年./-]\s*\d{0,2}\s*[月./-]?\s*\d{0,2}\s*日?\s*', txt):
            continue
        if re.search(r'(时间表|日程|进度安排|周计划)', txt) and len(txt) <= 60:
            continue

        # 表单类字段：保留标签，值改成 xxx（如「研究方向：xxx」）
        hit = next((k for k in KEEP_LABELS if k in txt), None)
        if hit:
            m = re.match(r'^([^：:]{1,10})[：:]\s*(.*)$', txt)
            if m:
                label = m.group(1).strip()
                if re.search(r'(日期|时间)', label):
                    new_txt = "%s：xxxx/xx/xx" % label
                else:
                    new_txt = "%s：xxx" % label
            else:
                new_txt = "%s：xxx" % txt.strip()
            set_text(sp, new_txt)
            stats['body'] += 1
            continue

        # 校名识别
        if re.search(r'(大学|学院|学校|University|College)', txt) and len(txt) <= 24:
            if set_text(sp, "学校名称"):
                stats['logo'] += 1
                continue

        _ab = abs_xfrm(sp)
        x3 = y3 = 999
        h3w3 = (0, 0)
        if _ab:
            x3, y3, w3, h3 = _ab
            if not (0 <= x3 <= 1280 and 0 <= y3 <= 720):
                x3 = y3 = 999
            h3w3 = (w3, h3)
        sz = max_font_sz(sp)

        w3, h3 = h3w3
        is_nav = (y3 < 130 and h3 <= 70) or (x3 < 160 and w3 <= 160 and h3 <= 200)
        if is_nav:                          # 顶部 / 左侧导航栏
            label = "标题"
        elif sz >= 3200:
            label = "标题"
        elif re.match(r'^[^：:]{1,8}[：:]\s*$', txt):          # 「短期目标：」→ 小标题
            label = "小标题"
        elif re.match(r'^[^：:]{1,8}[：:]\s*\S', txt) and len(txt) > 10:
            label = "小标题：正文"                              # 「短期目标：完成…」→ 小标题：正文
        elif (len(txt) <= 12 and not re.search(r'[。！？；]', txt)
              and (sz >= 2000 or ('b="1"' in etree.tostring(sp).decode('utf-8', 'ignore')))):
            label = "小标题"                                    # 短而加粗 = 小标题
        else:
            label = "正文"
        if set_text(sp, label):
            if label == "标题":
                stats['title'] += 1
            elif label.startswith("小标题"):
                stats.setdefault('subtitle', 0)
                stats['subtitle'] += 1
            else:
                stats['body'] += 1
    # 表格单元格文字 → 正文
    for tbl in root.iter('{%s}tbl' % A):
        for tc in tbl.iter('{%s}tc' % A):
            runs = [r for r in tc.iter('{%s}r' % A)]
            if not runs:
                continue
            txt = ''.join(t.text or '' for t in tc.iter('{%s}t' % A)).strip()
            if not txt:
                continue
            label = "小标题" if re.match(r'^[^：:]{1,8}[：:]\s*$', txt) else "正文"
            for r in runs[1:]:
                r.getparent().remove(r)
            t0 = runs[0].find('{%s}t' % A)
            if t0 is None:
                t0 = etree.SubElement(runs[0], '{%s}t' % A)
            t0.text = label
            stats['body'] += 1

    # 兜底：仍残留的"内容尺寸"图片一律换成「图片」占位
    for pic in list(root.iter('{%s}pic' % P)):
        b = abs_xfrm(pic)
        if not b:
            continue
        w, h = b[2], b[3]
        if w >= 1100 or w < 200 or h < 150:
            continue
        ph = make_placeholder(pic, "图片", box=b)
        if ph is None:
            continue
        parent = pic.getparent()
        parent.remove(pic)
        (spTree if spTree is not None else parent).append(ph)
        stats['pic'] += 1

    relabel_placeholders(root, _pic_boxes, _chart_boxes)
    return stats


def delete_first_slides(path, keep_from=4):
    """删除前 keep_from-1 页"""
    z = zipfile.ZipFile(path)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()

    pres = etree.fromstring(items['ppt/presentation.xml'])
    lst = pres.find('{%s}sldIdLst' % P)
    ids = lst.findall('{%s}sldId' % P)
    drop = set()
    for el in ids[:keep_from - 1]:
        rid = el.get('{%s}id' % R)
        drop.add(el)
        lst.remove(el)

    rels = etree.fromstring(items['ppt/_rels/presentation.xml.rels'])
    for rel in list(rels):
        if rel.get('Id') in {d.get('{%s}id' % R) for d in drop}:
            rels.remove(rel)
    items['ppt/_rels/presentation.xml.rels'] = etree.tostring(rels, xml_declaration=True,
                                                              encoding='UTF-8', standalone=True)

    # 重排序：sldIdLst 保持剩余顺序
    for i, el in enumerate(lst.findall('{%s}sldId' % P), 1):
        el.set('id', str(255 + i))
    items['ppt/presentation.xml'] = etree.tostring(pres, xml_declaration=True,
                                                   encoding='UTF-8', standalone=True)

    # 把被删的 slide XML 内容清空（保留部件结构，避免 rels 断裂）
    names = sorted([n for n in items if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)],
                   key=lambda s: int(re.findall(r'\d+', s)[0]))
    for n in names[:keep_from - 1]:
        root = etree.fromstring(items[n])
        sptree = root.find('.//{%s}spTree' % P)
        if sptree is not None:
            for ch in list(sptree):
                if etree.QName(ch).localname not in ('nvGrpSpPr', 'grpSpPr'):
                    sptree.remove(ch)
        items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return len(names) - (keep_from - 1)




# ---------- 校徽 / 校名 定点替换（坐标表来自 logo_boxes.json） ----------
BOXES_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logo_boxes.json')


def load_boxes(key):
    if not os.path.exists(BOXES_PATH):
        return None
    d = json.load(open(BOXES_PATH, encoding='utf-8'))
    return d.get(key)


def cover_box(root, box, label):
    """删除与 box 重叠的图形，并在该位置放白底 + 标签"""
    bx, by, bw, bh = box
    area = bw * bh
    spTree = root.find('.//{%s}spTree' % P)
    if spTree is None:
        return 0
    removed = 0
    for el in list(root.iter()):
        ln = etree.QName(el).localname
        if ln not in ('sp', 'pic', 'grpSp'):
            continue
        xf = el.find('.//{%s}xfrm' % A)
        if xf is None:
            continue
        off = xf.find('{%s}off' % A)
        ext = xf.find('{%s}ext' % A)
        if off is None or ext is None:
            continue
        x, y = int(off.get('x')) // EMU, int(off.get('y')) // EMU
        w, h = int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU
        ix = max(0, min(x + w, bx + bw) - max(x, bx))
        iy = max(0, min(y + h, by + bh) - max(y, by))
        same_scale = w <= bw * 1.6 and h <= bh * 1.6   # 只处理与原框相当大小的对象
        if same_scale and ix * iy >= area * 0.5:
            if el.getparent() is not None:
                el.getparent().remove(el)
                removed += 1
    ph = make_placeholder(None, label, box=box)
    if ph is not None:
        spTree.append(ph)
    return removed


def apply_logo_boxes(root, boxes):
    n = 0
    if not boxes:
        return n
    if boxes.get('logo'):
        n += cover_box(root, boxes['logo'], "校徽")
    if boxes.get('school'):
        n += cover_box(root, boxes['school'], "学校名称")
    return n

def process(src, dst, keep_from=4, boxes=None):
    shutil.copyfile(src, dst)
    rest = delete_first_slides(dst, keep_from)

    z = zipfile.ZipFile(dst)
    items = {n: z.read(n) for n in z.namelist()}
    z.close()

    total = dict(title=0, body=0, pic=0, logo=0, table=0)
    slides = sorted([n for n in items if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)],
                    key=lambda s: int(re.findall(r'\d+', s)[0]))
    # 母版与版式同样清理（校徽/校名常驻这里）
    extra = [n for n in items if re.match(r'ppt/(slideMasters/slideMaster\d+|slideLayouts/slideLayout\d+)\.xml$', n)]
    for i, n in enumerate(slides):
        root = etree.fromstring(items[n])
        st = scrub_slide(root)
        for k in total:
            total[k] += st[k]
        if boxes and i == 0:          # 仅封面做校徽/校名覆盖；内容页由母版统一替换
            apply_logo_boxes(root, boxes)
        else:                         # 内容页去掉「学校名称」占位
            for sp in list(root.iter('{%s}sp' % P)):
                t = ''.join(x.text or '' for x in sp.iter('{%s}t' % A)).strip()
                if t == '学校名称':
                    sp.getparent().remove(sp)
        items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    # 版式/母版只做校徽校名覆盖，保留原有导航栏、目录大矩形等版式元素
    for n in extra:
        root = etree.fromstring(items[n])
        for pic in list(root.iter('{%s}pic' % P)):
            xf = pic.find('.//{%s}xfrm' % A)
            if xf is None:
                continue
            off = xf.find('{%s}off' % A)
            ext = xf.find('{%s}ext' % A)
            if off is None or ext is None:
                continue
            w = int(ext.get('cx')) // EMU
            h = int(ext.get('cy')) // EMU
            if w >= 1100:
                continue
            ph = make_placeholder(pic, "校徽")
            if ph is None:
                continue
            parent = pic.getparent()
            idx = list(parent).index(pic)
            parent.remove(pic)
            parent.insert(idx, ph)
        items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for k, b in items.items():
            zo.writestr(k, b)
    return rest, total


if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    keep = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    key = sys.argv[4] if len(sys.argv) > 4 else None
    rest, total = process(src, dst, keep, load_boxes(key) if key else None)
    print("保留页数 %d | 标题 %d / 正文 %d / 图片位 %d / 校徽位 %d / 表格位 %d"
          % (rest, total['title'], total['body'], total['pic'], total['logo'], total['table']))
