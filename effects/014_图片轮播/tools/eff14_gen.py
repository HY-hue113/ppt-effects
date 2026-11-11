"""效果14 · 通用生成器：N 张图 → N 标签 / N+1 页
规则：P1 开场 C 位 = 图1+图2 重叠；Pk: A=图k-1 B=图k C=图k+1 D=图k-2（越界隐藏）
"""
import os, re, json, subprocess, sys, zipfile
from shutil import copyfile
from lxml import etree

SK = r'<USER_HOME>\<DSH_DIR>\skills\ppt-master'
PY = sys.executable
SRC = r'<LIB_ROOT>\effects\014_图片轮播\reference\效果14.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
CM = 360000
PAGE_W = 33.87

IMGS = [
    r'<DSH_ATTACHMENTS>\v1\objects\a8\a8eb323b40a77068b37401a8a34cfff3a6e9aff7542b01ef20f4a536e441ffd9',
    r'<DSH_ATTACHMENTS>\v1\objects\46\469e386a512252f3920a9fc9d1fbbaafcf0760a8d520952a700341f8926b2633',
    r'<DSH_ATTACHMENTS>\v1\objects\f5\f544eae007a239534e0e8bcee7c7de03e9d4122587d4977bae3b643d4b6c01e9',
    r'<DSH_ATTACHMENTS>\v1\objects\0d\0d3944e073c337e85bdf488e2650367d661045bc068c10524fbc8bcc485a7bf6',
]
N = len(IMGS)
PAGES = N + 1
OUT = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\out\效果14-四图.pptx'

PROJ = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\projects\eff14_gen4'
SVDF = os.path.join(PROJ, 'authoring-svg-flat')

r = subprocess.run([PY, os.path.join(SK, 'scripts', 'pptx_to_svg.py'), SRC, '-o', PROJ,
                    '--inheritance-mode', 'both', '--roundtrip'],
                   capture_output=True, text=True, encoding='utf-8', errors='ignore')
print('import rc', r.returncode)
if r.returncode != 0:
    print((r.stderr or '')[-500:]); sys.exit(1)

plan = [{'source_slide': i} for i in range(1, PAGES + 1)]
json.dump({'schema': 'ppt-master.roundtrip-page-plan.v1', 'pages': plan},
          open(os.path.join(PROJ, 'page_plan.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

r = subprocess.run([PY, os.path.join(SK, 'scripts', 'svg_to_pptx.py'), PROJ, '-o', OUT, '--roundtrip'],
                   capture_output=True, text=True, encoding='utf-8', errors='ignore')
print('export rc', r.returncode)
if r.returncode != 0:
    print((r.stdout or '')[-500:]); print((r.stderr or '')[-500:]); sys.exit(1)

z = zipfile.ZipFile(OUT)
items = {n: z.read(n) for n in z.namelist()}
z.close()

SLOT = PAGE_W / N
NUM_W = 1.9


def add_media(rels, rid_box, name, path):
    items[name] = open(path, 'rb').read()
    rid = 'rId%d' % rid_box[0]
    rid_box[0] += 1
    e = etree.SubElement(rels, '{%s}Relationship' % PR)
    e.set('Id', rid)
    e.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    e.set('Target', '../media/' + os.path.basename(name))
    return rid


for idx in range(1, PAGES + 1):
    sname = 'ppt/slides/slide%d.xml' % idx
    rels_name = 'ppt/slides/_rels/slide%d.xml.rels' % idx
    if sname not in items:
        continue
    root = etree.fromstring(items[sname])
    rels = etree.fromstring(items[rels_name])
    rid_box = [max(int(x.get('Id')[3:]) for x in rels) + 1]
    k = idx - 1

    # ---- 图片 ----
    def pick(*cands):
        for c in cands:
            if 1 <= c <= N:
                return c
        return None

    for sp in root.iter('{%s}pic' % P):
        spPr = sp.find('{%s}spPr' % P)
        xf = spPr.find('{%s}xfrm' % A) if spPr is not None else None
        off = xf.find('{%s}off' % A) if xf is not None else None
        ext = xf.find('{%s}ext' % A) if xf is not None else None
        blip = sp.find('{%s}blipFill/{%s}blip' % (P, A))
        if off is None or ext is None or blip is None:
            continue
        x = int(off.get('x')) / CM
        y = int(off.get('y')) / CM
        w = int(ext.get('cx')) / CM
        if x < -1:
            role = 'C'
        elif x >= 34:
            role = 'D'
        elif w >= 10:
            role = 'A'
        else:
            role = 'B'
        if k == 0:                       # 开场页：C 位放 图1/图2
            seq = 1 if (role == 'C' and x < -1) else None
            if seq is None:
                off.set('y', str(int(-50 * CM)))
                continue
            # 两个 C 位分别放 图1、图2
            same = [s for s in root.iter('{%s}pic' % P)
                    if (s.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A)) is not None
                        and int(s.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A)).get('x')) < -1 * CM)]
            order = same.index(sp) if sp in same else 0
            seq = order + 1 if order < 2 else None
            if seq is None:
                off.set('y', str(int(-50 * CM)))
                continue
            rid = add_media(rels, rid_box, 'ppt/media/e14g_%d_%d.jpg' % (idx, rid_box[0]), IMGS[seq - 1])
            blip.set('{%s}embed' % R, rid)
            print('p%d %s -> 图%d' % (idx, role, seq))
            continue
        seq = {'A': pick(k), 'B': pick(k + 1), 'C': pick(k + 2), 'D': pick(k - 1)}[role]
        if seq is None:
            off.set('y', str(int(-50 * CM)))
        else:
            rid = add_media(rels, rid_box, 'ppt/media/e14g_%d_%d.jpg' % (idx, rid_box[0]), IMGS[seq - 1])
            blip.set('{%s}embed' % R, rid)
        print('p%d %s -> %s' % (idx, role, ('图%d' % seq) if seq else '隐藏'))

    # ---- 标签 ----
    for sp in root.iter('{%s}sp' % P):
        ts = [t.text.strip() for t in sp.iter('{%s}t' % A) if t.text and t.text.strip()]
        if len(ts) != 1 or not re.fullmatch(r'\d{2}', ts[0]):
            continue
        n = int(ts[0])
        off = sp.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A))
        ext = sp.find('{%s}spPr/{%s}xfrm/{%s}ext' % (P, A, A))
        if off is None or ext is None:
            continue
        if n > N:
            off.set('y', str(int(-20 * CM)))
        else:
            w = int(ext.get('cx'))
            center = SLOT * (n - 1) + SLOT / 2
            off.set('x', str(int(center * CM - w / 2)))
            off.set('y', str(int(1.4 * CM)))

    # ---- 高亮块 ----
    num = None
    for sp in root.iter('{%s}sp' % P):
        ts = [t.text.strip() for t in sp.iter('{%s}t' % A) if t.text and t.text.strip()]
        if len(ts) == 1 and ts[0] == '%02d' % k and k >= 1:
            num = sp
            break
    for sp in root.iter('{%s}sp' % P):
        c = sp.find('.//{%s}cNvPr' % P)
        if c is None or not c.get('name', '').startswith('矩形: 圆角'):
            continue
        off = sp.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A))
        ext = sp.find('{%s}spPr/{%s}xfrm/{%s}ext' % (P, A, A))
        if off is None or ext is None:
            continue
        if num is not None:
            noff = num.find('{%s}spPr/{%s}xfrm/{%s}off' % (P, A, A))
            next_ = num.find('{%s}spPr/{%s}xfrm/{%s}ext' % (P, A, A))
            nx, ny = int(noff.get('x')), int(noff.get('y'))
            nw, nh = int(next_.get('cx')), int(next_.get('cy'))
            hw, hh = int(ext.get('cx')), int(ext.get('cy'))
            off.set('x', str(int(nx + nw / 2 - hw / 2)))
            off.set('y', str(int(ny + nh / 2 - hh / 2)))
        else:
            off.set('x', str(int(-6 * CM)))
            off.set('y', str(int(0.93 * CM)))

    # ---- 背景 ----
    bg_seq = k if k >= 1 else 1
    media = 'ppt/media/e14gb_%d.jpg' % idx
    rid_bg = add_media(rels, rid_box, media, IMGS[bg_seq - 1])
    bg = root.find('{%s}cSld/{%s}bg' % (P, P))
    if bg is not None:
        blip = bg.find('.//{%s}blip' % A)
        if blip is not None:
            blip.set('{%s}embed' % R, rid_bg)

    items[sname] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    items[rels_name] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zo:
    for n, b in items.items():
        zo.writestr(n, b)
print('saved', OUT, round(os.path.getsize(OUT) / 1024), 'KB')
