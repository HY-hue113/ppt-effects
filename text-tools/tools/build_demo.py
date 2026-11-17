# -*- coding: utf-8 -*-
"""文本底卡处理区 · 示范稿（6 页）"""
import io, os, sys, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from card_lib import card, chip, tab, badge

SK = r"<USER_HOME>\<DSH_DIR>\skills\ppt-master"
PY = sys.executable
WS = r"<USER_HOME>\Documents\<DSH_HOME>\default-workspace\projects\texttools_demo"
SVG, IMG = os.path.join(WS, "svg_output"), os.path.join(WS, "images")
OUT = r"<LIB_ROOT>\text-tools\demo\reference\文本底卡_示范.pptx"
for d in (SVG, IMG):
    os.makedirs(d, exist_ok=True)
    for f in os.listdir(d):
        os.remove(os.path.join(d, f))
os.makedirs(os.path.dirname(OUT), exist_ok=True)

W, H = 1280, 720
BG, INK, GOLD, GREY = "#F6F4EE", "#1F2A2E", "#C3AE53", "#8A8578"


def text(x, y, s, size=20, color=INK, bold=False, anchor="start"):
    return ('<text x="%g" y="%g" font-size="%g" fill="%s" font-family="Microsoft YaHei"%s text-anchor="%s">%s</text>'
            % (x, y, size, color, ' font-weight="bold"' if bold else '', anchor, s))


def page(title, sub, parts):
    head = text(60, 62, title, 30, INK, True) + text(60, 94, sub, 16, GREY)
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d" lang="zh-CN">\n'
            '<rect x="0" y="0" width="%d" height="%d" fill="%s"/>\n%s\n%s\n</svg>'
            % (W, H, W, H, W, H, BG, head, "\n".join(parts)))


pages = []

# P1 四种底卡
parts = []
styles = [("plain", "plain 白底"), ("tint", "tint 浅蓝"), ("bar", "bar 金条"), ("line", "line 线框")]
for i, (st, label) in enumerate(styles):
    x = 70 + i * 300
    f = card(os.path.join(IMG, "c_%s.png" % st), 260, 150, style=st)
    parts.append('<image href="../images/%s" x="%d" y="190" width="260" height="150"/>' % (os.path.basename(f), x))
    parts.append(text(x, 370, label, 17, GREY))
pages.append(page("四种底卡", "plain / tint / bar / line —— 同一尺寸 260×150", parts))

# P2 小标题标签 + 正文卡（分开）
parts = []
t = tab(os.path.join(IMG, "t1.png"), 240, 46)
parts.append('<image href="../images/%s" x="80" y="170" width="240" height="46"/>' % os.path.basename(t))
parts.append(text(200, 202, "小标题", 22, "#FFFFFF", True, "middle"))
k = card(os.path.join(IMG, "k1.png"), 600, 220, style="bar")
parts.append('<image href="../images/%s" x="80" y="240" width="600" height="220"/>' % os.path.basename(k))
for j, ln in enumerate(["正文与标题分开承托", "正文用 bar 卡，纵向留白充足", "标题单独成标签，层级清楚"]):
    parts.append(text(120, 300 + j * 44, "· " + ln, 19, "#33404D"))
pages.append(page("标题与正文分开承托", "tab 标签 + bar 正文卡", parts))

# P3 要点胶囊
parts = []
for j, ln in enumerate(["门址考：以方志与旧地图复原十门", "口述史：十门十声田野采集", "空间观察：三类地块定点记录", "话语分析：异说表逐条对照"]):
    c = chip(os.path.join(IMG, "ch%d.png" % j), 620, 48)
    parts.append('<image href="../images/%s" x="80" y="%d" width="620" height="48"/>' % (os.path.basename(c), 180 + j * 70))
    parts.append(text(112, 212 + j * 70, ln, 19, "#33404D"))
pages.append(page("要点胶囊：一条一卡", "chip 纵向排列，适合 3–4 条要点", parts))

# P4 2×2 网格
parts = []
for j in range(4):
    cx = 80 + (j % 2) * 320
    cy = 180 + (j // 2) * 200
    k = card(os.path.join(IMG, "g%d.png" % j), 300, 180, style="plain")
    parts.append('<image href="../images/%s" x="%d" y="%d" width="300" height="180"/>' % (os.path.basename(k), cx, cy))
    b = badge(os.path.join(IMG, "b%d.png" % j), 44)
    parts.append('<image href="../images/%s" x="%d" y="%d" width="44" height="44"/>' % (os.path.basename(b), cx + 20, cy + 20))
    parts.append(text(cx + 42, cy + 50, "0%d" % (j + 1), 18, "#FFFFFF", True, "middle"))
    parts.append(text(cx + 20, cy + 100, "要点标题 %d" % (j + 1), 20, INK, True))
    parts.append(text(cx + 20, cy + 138, "两行以内的说明文字", 16, GREY))
pages.append(page("2×2 网格卡片", "每格一张 plain 卡 + 序号块，适合并列要点", parts))

# P5 叠在图上（line 卡）
parts = ['<rect x="80" y="170" width="640" height="360" rx="10" fill="#DCE8F6"/>']
k = card(os.path.join(IMG, "l1.png"), 420, 200, style="line")
parts.append('<image href="../images/%s" x="120" y="420" width="420" height="200"/>' % os.path.basename(k))
parts.append(text(150, 470, "line 卡片叠在图上", 20, "#FFFFFF", True))
parts.append(text(150, 505, "只有细边不遮图，适合压在照片下缘", 16, "#FFFFFF"))
pages.append(page("叠图用法", "line 款：无填充，只有细边", parts))

# P6 混排
parts = []
t = tab(os.path.join(IMG, "t2.png"), 200, 42)
parts.append('<image href="../images/%s" x="80" y="160" width="200" height="42"/>' % os.path.basename(t))
parts.append(text(180, 189, "混排示例", 20, "#FFFFFF", True, "middle"))
for j, (txt, st) in enumerate([("一条要点", "chip"), ("两条要点", "chip")]):
    c = chip(os.path.join(IMG, "mc%d.png" % j), 380, 44)
    parts.append('<image href="../images/%s" x="80" y="%d" width="380" height="44"/>' % (os.path.basename(c), 230 + j * 62))
    parts.append(text(112, 259 + j * 62, txt, 18, "#33404D"))
k = card(os.path.join(IMG, "mk.png"), 420, 160, style="tint")
parts.append('<image href="../images/%s" x="500" y="230" width="420" height="160"/>' % os.path.basename(k))
parts.append(text(520, 285, "右侧强调卡", 20, INK, True))
parts.append(text(520, 320, "tint 底 + 两行说明", 16, GREY))
pages.append(page("混排：标签 + 胶囊 + 强调卡", "同一页里三种容器各司其职", parts))

# P7 时间轴胶囊
parts = []
parts.append('<rect x="104" y="190" width="3" height="400" fill="#D8DFC9"/>')
for j, (when, what) in enumerate([("2026.11", "文献研读与门址考据"), ("2027.01", "实地踏勘与影像采集"),
                                  ("2027.03", "问卷发放与定点观察"), ("2027.04", "数字地图与成果展")]):
    yy = 200 + j * 100
    parts.append('<circle cx="105" cy="%d" r="7" fill="#C3AE53"/>' % (yy - 6))
    c = chip(os.path.join(IMG, "tl%d.png" % j), 420, 46)
    parts.append('<image href="../images/%s" x="140" y="%d" width="420" height="46"/>' % (os.path.basename(c), yy - 30))
    parts.append(text(170, yy + 4, when, 16, "#8A8578"))
    parts.append(text(280, yy + 4, what, 18, "#33404D"))
pages.append(page("时间轴胶囊", "chip 竖排 + 细线 + 圆点，适合时间安排", parts))

# P8 引文卡
parts = []
k = card(os.path.join(IMG, "q1.png"), 720, 260, style="tint")
parts.append('<image href="../images/%s" x="100" y="180" width="720" height="260"/>' % os.path.basename(k))
parts.append(text(130, 250, "“", 64, "#C3AE53", True))
parts.append(text(180, 270, "城的实体退场了，市井的记忆却没有散去——", 21, "#1F2A2E"))
parts.append(text(180, 310, "它留在老人的讲述里，留在菜场的方言里。", 21, "#1F2A2E"))
parts.append(text(180, 380, "—— 受访者 03，2027.01 于凤山水城门", 15, "#8A8578"))
pages.append(page("引文卡", "tint 底卡 + 引号 + 出处，访谈摘录专用", parts))

# P9 数据块
parts = []
for j, (num, label) in enumerate([("500", "直接接触人数"), ("3", "Citywalk 线路"), ("1890.60", "预算合计（元）")]):
    x = 80 + j * 390
    k = card(os.path.join(IMG, "d%d.png" % j), 350, 240, style="plain")
    parts.append('<image href="../images/%s" x="%d" y="180" width="350" height="240"/>' % (os.path.basename(k), x))
    fs = 56 if len(num) <= 3 else 40
    parts.append(text(x + 175, 290, num, fs, "#1F4E79", True, "middle"))
    parts.append(text(x + 175, 345, label, 17, "#8A8578", False, "middle"))
    parts.append('<rect x="%d" y="%d" width="60" height="3" fill="#C3AE53"/>' % (x + 145, 365))
pages.append(page("数据强调块", "大号数字 + 说明 + 金线，适合预算/人数/场次", parts))


# P10 方卡三种分型
parts = []
for i, (st, label) in enumerate([("corner", "corner 折角"), ("duo", "duo 双线"), ("topbar", "topbar 顶条")]):
    x = 70 + i * 380
    f = card(os.path.join(IMG, "split_%s.png" % st), 330, 190, style=st)
    parts.append('<image href="../images/%s" x="%d" y="180" width="330" height="190"/>' % (os.path.basename(f), x))
    parts.append(text(x, 400, label, 17, GREY))
    parts.append(text(x + 16, 250, "方卡分型", 20, INK, True))
    parts.append(text(x + 16, 288, "同主类型，改装饰", 15, GREY))
pages.append(page("方卡三种分型", "corner / duo / topbar —— 只换装饰，不换主类型", parts))

for i, s in enumerate(pages, 1):
    open(os.path.join(SVG, "slide_%02d.svg" % i), "w", encoding="utf-8").write(s)
print("SVG", len(pages), "页")

subprocess.run([PY, os.path.join(SK, "scripts", "svg_quality_checker.py"), WS,
                "--quick-generate", "--canonical-authoring", "--stage", "final", "--json"], capture_output=True)
r = subprocess.run([PY, os.path.join(SK, "scripts", "svg_to_pptx.py"), WS, "-f", "ppt169",
                    "--quick-generate", "-t", "none", "-o", OUT], capture_output=True, text=True, encoding="utf-8", errors="ignore")
print("export rc", r.returncode)
if r.returncode != 0:
    print((r.stdout or "")[-500:]); print((r.stderr or "")[-500:])
else:
    print("saved", OUT, round(os.path.getsize(OUT) / 1024), "KB")

# WIDEN-PATCH：演示页底部的分型标签加宽文本框，避免折行
import zipfile as _zf
from lxml import etree as _et

_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
if os.path.exists(OUT):
    _z = _zf.ZipFile(OUT)
    _items = {n: _z.read(n) for n in _z.namelist()}
    _z.close()
    _name = 'ppt/slides/slide10.xml'
    if _name in _items:
        _root = _et.fromstring(_items[_name])
        _n = 0
        for _sp in _root.iter('{%s}sp' % _P):
            _off = _sp.find('{%s}spPr/{%s}xfrm/{%s}off' % (_P, _A, _A))
            _ext = _sp.find('{%s}spPr/{%s}xfrm/{%s}ext' % (_P, _A, _A))
            if _off is None or _ext is None:
                continue
            if int(_off.get('y')) / 9525 < 380:
                continue
            if _sp.find('.//{%s}t' % _A) is None:
                continue
            _body = _sp.find('.//{%s}bodyPr' % _A)
            if _body is not None:
                _body.set('wrap', 'none')
            _cx = int(_ext.get('cx'))
            _ext.set('cx', str(int(_cx * 1.6)))
            _n += 1
        if _n:
            _items[_name] = _et.tostring(_root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True)
            with _zf.ZipFile(OUT, 'w', _zf.ZIP_DEFLATED) as _zo:
                for _k, _b in _items.items():
                    _zo.writestr(_k, _b)
            print('加宽分型标签文本框:', _n, '个')
