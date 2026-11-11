"""效果12 · 中式滑转（严格对齐用户手改稿）

页数 = 1 + N。第 i 页（i≥2）同时保留「上一张」与「当前张」两块图片板/色板/文字组，
两页之间靠 x 位移滑动；方向一左一右交替。
"""
import os, re, argparse, zipfile
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'

CM = 360000
PAGE_W, PAGE_H = 33.87, 19.05
PANEL_W = 15.37
BODY_W = 19.62
BODY_X = 7.21

X_PIC_R, X_PIC_L = 22.91, -4.14
X_SOL_R, X_SOL_L = 22.81, -4.24
X_TXT_R, X_TXT_L = 29.05, 1.77
LAG = 0.36                     # 上一张相对当前张的错位

PAPER = 'F3EADC'
GOLD = 'C8A974'
ADJ = int(35 / 90 * 100000)
BG_SEQ = ['3A1113', '3A1113', 'C35D09', 'C35D09', '1F3A5F', '1F3A5F', '2E4A3A']


def emu(cm):
    return Emu(int(cm * CM))


def quad(sh, adj=ADJ):
    spPr = sh._element.spPr
    g = spPr.find('{%s}prstGeom' % A)
    if g is None:
        g = etree.SubElement(spPr, '{%s}prstGeom' % A)
    g.set('prst', 'parallelogram')
    for c in list(g):
        g.remove(c)
    av = etree.SubElement(g, '{%s}avLst' % A)
    gd = etree.SubElement(av, '{%s}gd' % A)
    gd.set('name', 'adj')
    gd.set('fmla', 'val %d' % adj)


def pic_fill(sh, rid):
    spPr = sh._element.spPr
    for tag in ('solidFill', 'gradFill', 'noFill', 'blipFill'):
        for e in spPr.findall('{%s}%s' % (A, tag)):
            spPr.remove(e)
    bf = etree.SubElement(spPr, '{%s}blipFill' % A)
    bf.set('dpi', '0')
    bf.set('rotWithShape', '1')
    b = etree.SubElement(bf, '{%s}blip' % A)
    b.set('{%s}embed' % R, rid)
    st = etree.SubElement(bf, '{%s}stretch' % A)
    etree.SubElement(st, '{%s}fillRect' % A)


class Builder:
    def __init__(self, titles, bodies, t_font='华文行楷', b_font='微软雅黑',
                 t_size=40, b_size=16, show_body=True):
        self.titles, self.bodies = titles, bodies
        self.t_font, self.b_font = t_font, b_font
        self.t_size, self.b_size = t_size, b_size
        self.show_body = show_body
        self.prs = Presentation()
        self.prs.slide_width = emu(PAGE_W)
        self.prs.slide_height = emu(PAGE_H)
        self.blank = prs_layout = self.prs.slide_layouts[6]

    def panel(self, s, x, w, rid=None, color=None, name='面板', line=True):
        sh = s.shapes.add_shape(1, emu(x), Emu(0), emu(w), emu(PAGE_H))
        quad(sh)
        if rid:
            pic_fill(sh, rid)
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = RGBColor.from_string(color)
            if line:
                sh.line.color.rgb = RGBColor.from_string(GOLD)
                sh.line.width = Pt(1)
            else:
                sh.line.fill.background()
        sh.name = name
        return sh

    def text_group(self, s, x, tag, page_i=1):
        g = s.shapes.add_group_shape()
        title = self.titles[min(page_i - 1, len(self.titles) - 1)]
        body = self.bodies[min(page_i - 1, len(self.bodies) - 1)]
        t1 = s.shapes.add_textbox(emu(x + 0.45), emu(4.2), emu(2.6), emu(6.0))
        f1 = t1.text_frame
        f1.word_wrap = True
        f1._txBody.bodyPr.set('vert', 'eaVert')
        r1 = f1.paragraphs[0].add_run()
        r1.text = title
        r1.font.size = Pt(self.t_size)
        r1.font.name = self.t_font
        r1.font.color.rgb = RGBColor.from_string(PAPER)
        items = [t1]
        if self.show_body and body:
            t2 = s.shapes.add_textbox(emu(x + 0.45), emu(11.2), emu(3.0), emu(1.2))
            r2 = t2.text_frame.paragraphs[0].add_run()
            r2.text = body
            r2.font.size = Pt(self.b_size)
            r2.font.name = self.b_font
            r2.font.color.rgb = RGBColor.from_string(PAPER)
            items.append(t2)
        for t in items:
            g.shapes._spTree.append(t._element)
        g.name = '文字组%s' % tag
        return g

    def rid_for(self, s, img):
        tmp = s.shapes.add_picture(img, 0, 0)
        rid = tmp._element.blipFill.blip.rEmbed
        s.shapes._spTree.remove(tmp._element)
        return rid

    def page(self, i, images):
        """i: 1..N+1"""
        s = self.prs.slides.add_slide(self.blank)
        bg = BG_SEQ[min((i - 1) // 2, len(BG_SEQ) - 1)]
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(bg)

        if i == 1:
            # 封面：满幅色板 + 第 1 张图片板 + 正文板 + 文字组
            self.panel(s, -4.94, 34.36, color=bg, name='色板1', line=False)
            self.panel(s, 9.34, PANEL_W, rid=self.rid_for(s, images[0]), name='图片板1')
            self.panel(s, BODY_X, BODY_W, color=PAPER, name='正文板', line=False)
            self.text_group(s, 13.6, 1, 1)
            return s

        right = (i % 2 == 0)          # 偶数页在右、奇数页在左（对齐手改稿 p2 右 / p3 左 / p4 右）
        pic_x = X_PIC_R if right else X_PIC_L
        sol_x = X_SOL_L if right else X_SOL_R
        txt_x = X_TXT_L if right else X_TXT_R
        sign = 1 if right else -1

        cur = min(i, len(images))     # 当前张
        prev = max(1, cur - 1)        # 上一张
        # 上下两张叠放（上一张向页面中心错位 0.36cm）
        if prev != cur:
            self.panel(s, pic_x - sign * LAG, PANEL_W, rid=self.rid_for(s, images[prev - 1]),
                       name='图片板%d' % prev)
            self.panel(s, sol_x - sign * LAG, PANEL_W, color=BG_SEQ[min((prev - 1) // 2, len(BG_SEQ) - 1)],
                       name='色板%d' % prev)
        self.panel(s, pic_x, PANEL_W, rid=self.rid_for(s, images[cur - 1]), name='图片板%d' % cur)
        self.panel(s, sol_x, PANEL_W, color=bg, name='色板%d' % cur)
        self.panel(s, BODY_X, BODY_W, color=PAPER, name='正文板', line=False)
        if prev != cur:
            self.text_group(s, txt_x - sign * LAG, prev, prev)
        self.text_group(s, txt_x, cur, cur)
        return s

    def build(self, images):
        for i in range(1, len(images) + 2):
            self.page(i, images)
        return self.prs

    def save(self, out):
        self.prs.save(out)
        self._morph(out)

    def _morph(self, out):
        TMPL = ('<mc:AlternateContent xmlns:mc="' + MC + '" xmlns:p="' + P + '" xmlns:p14="' + P14 +
                '" xmlns:p159="' + P159 + '"><mc:Choice Requires="p159">'
                '<p:transition spd="slow" p14:dur="1500"><p159:morph option="byObject"/></p:transition>'
                '</mc:Choice><mc:Fallback><p:transition spd="slow" p14:dur="1500">'
                '<p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
        with zipfile.ZipFile(out) as z:
            items = {n: z.read(n) for n in z.namelist()}
        for n in list(items):
            m = re.match(r'ppt/slides/slide(\d+)\.xml$', n)
            if not m or int(m.group(1)) == 1:
                continue
            root = etree.fromstring(items[n])
            if root.find('{%s}transition' % P) is not None:
                continue
            root.append(etree.fromstring(TMPL.encode()))
            kids = list(root)
            order = {'cSld': 0, 'clrMapOvr': 1, 'AlternateContent': 2, 'timing': 3, 'extLst': 4}
            kids.sort(key=lambda el: order.get(etree.QName(el).localname, 9))
            for el in list(root):
                root.remove(el)
            for el in kids:
                root.append(el)
            items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for k, v in items.items():
                z.writestr(k, v)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', required=True)
    ap.add_argument('--title', action='append', help='每页标题，可重复：--title 第一章 --title 第二章')
    ap.add_argument('--body', action='append', help='每页正文，可重复')
    ap.add_argument('--title-font', default='华文行楷')
    ap.add_argument('--body-font', default='微软雅黑')
    ap.add_argument('--title-size', type=float, default=40)
    ap.add_argument('--body-size', type=float, default=16)
    ap.add_argument('--no-body', action='store_true', help='每页只放标题，不放正文')
    ap.add_argument('images', nargs='+')
    a = ap.parse_args()
    titles = a.title or ['标题']
    bodies = a.body or ['正文']
    b = Builder(titles, bodies, a.title_font, a.body_font,
                a.title_size, a.body_size, show_body=not a.no_body)
    b.build(a.images)
    b.save(a.out)
    print('页数', 1 + len(a.images), '| saved', a.out, round(os.path.getsize(a.out) / 1024), 'KB')
