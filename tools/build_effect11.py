"""效果11 · 中式滑转（按用户手改稿逐项对齐）

结构（自下而上）：
  i=1 : 图片板(9.34) → 米色正文板(7.21) → 满幅深色板(-4.94,34.36) → 文字
  i≥2 : 图片板(当前) → 图片板(上一张,-0.36) → 米色正文板 → 色板(当前)
        → 文字(当前) → 色板(上一张) → 文字(上一张)
文字位置固定：竖排标题 x=2.12(w2.54,h5.90)，正文 x=14.70(w5.5~6.8,h3.08)
面板左右交替：i 偶 → 图片在右(22.91)/色板在左(-4.24)；i 奇 → 反之
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
BODY_X, BODY_W = 7.21, 19.62
X_PIC_R, X_PIC_L = 22.91, -4.14
X_SOL_R, X_SOL_L = 22.81, -4.24
X_TITLE, W_TITLE, H_TITLE = 2.12, 2.54, 5.90
X_TEXT, W_TEXT, H_TEXT = 14.70, 6.00, 3.08
Y_TITLE, Y_TEXT = 4.73, 10.93
LAG = 0.36
PAPER = 'F3EADC'
INK = '2B2B2B'      # 正文落在米色板上，必须用深色
ADJ = int(35 / 90 * 100000)
BG_SEQ = ['3A1113', '3A1113', 'C35D09', 'C35D09', '1F3A5F', '1F3A5F']


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
        self.blank = self.prs.slide_layouts[6]

    def panel(self, s, x, w, rid=None, color=None, name='面板', line=False):
        sh = s.shapes.add_shape(1, emu(x), Emu(0), emu(w), emu(PAGE_H))
        quad(sh)
        if rid:
            pic_fill(sh, rid)
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = RGBColor.from_string(color)
            if line:
                sh.line.color.rgb = RGBColor.from_string(PAPER)
                sh.line.width = Pt(1)
            else:
                sh.line.fill.background()
        sh.name = name
        return sh

    def texts(self, s, page_i):
        title = self.titles[min(page_i - 1, len(self.titles) - 1)]
        body = self.bodies[min(page_i - 1, len(self.bodies) - 1)]
        t1 = s.shapes.add_textbox(emu(X_TITLE), emu(Y_TITLE), emu(W_TITLE), emu(H_TITLE))
        f1 = t1.text_frame
        f1.word_wrap = True
        f1._txBody.bodyPr.set('vert', 'eaVert')
        r1 = f1.paragraphs[0].add_run()
        r1.text = title
        r1.font.size = Pt(self.t_size)
        r1.font.name = self.t_font
        r1.font.color.rgb = RGBColor.from_string(PAPER)
        t1.name = '标题%d' % page_i
        if self.show_body and body:
            t2 = s.shapes.add_textbox(emu(X_TEXT), emu(Y_TEXT), emu(W_TEXT), emu(H_TEXT))
            f2 = t2.text_frame
            f2.word_wrap = True
            r2 = f2.paragraphs[0].add_run()
            r2.text = body
            r2.font.size = Pt(self.b_size)
            r2.font.name = self.b_font
            r2.font.color.rgb = RGBColor.from_string(INK)
            t2.name = '正文%d' % page_i
        return t1

    def rid_for(self, s, img):
        tmp = s.shapes.add_picture(img, 0, 0)
        rid = tmp._element.blipFill.blip.rEmbed
        s.shapes._spTree.remove(tmp._element)
        return rid

    def page(self, i, images):
        s = self.prs.slides.add_slide(self.blank)
        bg = BG_SEQ[min((i - 1) // 2, len(BG_SEQ) - 1)]
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(bg)

        if i == 1:
            self.panel(s, 9.34, PANEL_W, rid=self.rid_for(s, images[0]), name='图片板1')
            self.panel(s, BODY_X, BODY_W, color=PAPER, name='正文板')
            self.panel(s, -4.94, 34.36, color=bg, name='色板1')     # 满幅深色板压在最上
            ln = s.shapes.add_connector(1, emu(24.70), Emu(0), emu(29.42), emu(PAGE_H))
            ln.line.color.rgb = RGBColor.from_string(PAPER)
            ln.line.width = Pt(1)
            ln.name = '装饰线1'
            self.texts(s, 1)
            return s

        right = (i % 2 == 0)
        pic_x = X_PIC_R if right else X_PIC_L
        sol_x = X_SOL_L if right else X_SOL_R
        sign = 1 if right else -1
        cur = min(i, len(images))
        prev = max(1, cur - 1)

        if prev != cur:
            self.panel(s, pic_x - sign * LAG, PANEL_W, rid=self.rid_for(s, images[prev - 1]),
                       name='图片板%d' % prev)
        self.panel(s, pic_x, PANEL_W, rid=self.rid_for(s, images[cur - 1]), name='图片板%d' % cur)
        self.panel(s, BODY_X, BODY_W, color=PAPER, name='正文板')
        self.panel(s, sol_x, PANEL_W, color=bg, name='色板%d' % cur)
        self.texts(s, cur)
        if prev != cur:
            self.panel(s, sol_x - sign * LAG, PANEL_W, color=BG_SEQ[min((prev - 1) // 2, len(BG_SEQ) - 1)],
                       name='色板%d' % prev)
            self.texts(s, prev)
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
    ap.add_argument('--title', action='append')
    ap.add_argument('--body', action='append')
    ap.add_argument('--title-font', default='华文行楷')
    ap.add_argument('--body-font', default='微软雅黑')
    ap.add_argument('--title-size', type=float, default=40)
    ap.add_argument('--body-size', type=float, default=16)
    ap.add_argument('--no-body', action='store_true')
    ap.add_argument('images', nargs='+')
    a = ap.parse_args()
    b = Builder(a.title or ['标题'], a.body or ['正文'], a.title_font, a.body_font,
                a.title_size, a.body_size, show_body=not a.no_body)
    b.build(a.images)
    b.save(a.out)
    print('页数', 1 + len(a.images), '| saved', a.out, round(os.path.getsize(a.out) / 1024), 'KB')
