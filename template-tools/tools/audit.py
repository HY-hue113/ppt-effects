# -*- coding: utf-8 -*-
"""模板质量审计：未占位文本 / 占位漂移 / 图片漏占位

用法：python audit.py <源.pptx> <生成.pptx>
"""
import re
import sys
import zipfile

from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525

LABELS = {'标题', '正文', '小标题', '小标题：正文', '图片', '图表', '表格', '校徽', '学校名称', '目录'}
KEEP_PAT = re.compile(r'(指导教师|指导老师|导师|答辩人|答辩日期|日期|课题小组|研究方向|专业|姓名|学号|院系|班级|电话|邮箱|成员|汇报人|作者|单位|职称|大学|学院|学校)')
TIME_PAT = re.compile(r'(\d{4}\s*[年./-]|周|时间|时段|时间表|日程|进度)')


def parts(f):
    z = zipfile.ZipFile(f)
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
    m = {r.get('Id'): r.get('Target') for r in rels}
    order = ['ppt/' + m[s.get('{%s}id' % R)].replace('../', '') for s in pres.find('{%s}sldIdLst' % P)]
    return z, order


def box(el):
    xf = el.find('.//{%s}xfrm' % A)
    if xf is None:
        xf = el.find('.//{%s}xfrm' % P)
    if xf is None:
        return None
    off = xf.find('{%s}off' % A) or xf.find('{%s}off' % P)
    ext = xf.find('{%s}ext' % A) or xf.find('{%s}ext' % P)
    if off is None or ext is None:
        return None
    return (int(off.get('x')) // EMU, int(off.get('y')) // EMU,
            int(ext.get('cx')) // EMU, int(ext.get('cy')) // EMU)


def audit(src, gen):
    zs, ps = parts(src)
    zg, pg = parts(gen)
    issues = []
    for idx, (sp, gp) in enumerate(zip(ps, pg), 1):
        if sp not in zs.namelist() or gp not in zg.namelist():
            continue
        rs = etree.fromstring(zs.read(sp))
        rg = etree.fromstring(zg.read(gp))

        # 1) 未占位文本
        for s in rg.iter('{%s}sp' % P):
            t = ''.join(x.text or '' for x in s.iter('{%s}t' % A)).strip()
            if not t or t in LABELS or KEEP_PAT.search(t) or TIME_PAT.search(t):
                continue
            issues.append((idx, '未占位文本', t[:24]))

        # 2) 占位漂移：生成的占位框要能对上原始图片/图表位置
        src_boxes = []
        for el in rs.iter():
            if etree.QName(el).localname in ('pic', 'graphicFrame'):
                b = box(el)
                if b:
                    src_boxes.append(b)
        for s in rg.iter('{%s}sp' % P):
            t = ''.join(x.text or '' for x in s.iter('{%s}t' % A)).strip()
            if t not in ('图片', '图表', '表格'):
                continue
            b = box(s)
            if not b:
                continue
            ok = any(abs(b[0] - c[0]) <= 40 and abs(b[1] - c[1]) <= 40 for c in src_boxes)
            if not ok:
                issues.append((idx, '占位漂移', '%s @%d,%d' % (t, b[0], b[1])))

        # 3) 漏占位：原始内容图（较大）没有对应占位
        src_pics = []
        for el in rs.iter('{%s}pic' % P):
            b = box(el)
            if b and b[2] >= 200 and b[3] >= 150 and b[2] < 1100:
                src_pics.append(b)
        gen_boxes = []
        for s in rg.iter('{%s}sp' % P):
            t = ''.join(x.text or '' for x in s.iter('{%s}t' % A)).strip()
            if t == '图片':
                b = box(s)
                if b:
                    gen_boxes.append(b)
        for b in src_pics:
            ok = any(abs(b[0] - c[0]) <= 40 and abs(b[1] - c[1]) <= 40 for c in gen_boxes)
            if not ok:
                issues.append((idx, '漏占位', '@%d,%d %dx%d' % b))
    return issues


if __name__ == '__main__':
    iss = audit(sys.argv[1], sys.argv[2])
    from collections import Counter
    c = Counter(k for _, k, _ in iss)
    print('问题统计:', dict(c), '合计', len(iss))
    for it in iss[:40]:
        print('  p%-3d %-6s %s' % it)
