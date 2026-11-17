# -*- coding: utf-8 -*-
"""图片处理区 · 示范稿合并构建（10 页 = 原 6 页 + 新 4 页）

以后更新示范稿一律走这个脚本：旧页从 OLD_SVG 拷入，新页在 EXTRA 里写，最后一次性导出。
"""
import io, os, sys, shutil, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from mask_lib import shaped_image, outline_svg

OLD_WS = r"<ABS_PATH>"
WS = r"<USER_HOME>\Documents\<DSH_HOME>\default-workspace\projects\imagetools_demo10"
SVG = os.path.join(WS, "svg_output")
IMG = os.path.join(WS, "images")
SK = r"<USER_HOME>\<DSH_DIR>\skills\ppt-master"
PY = sys.executable
OUT = r"<USER_HOME>\Documents\<DSH_HOME>\default-workspace\out\图片取景形状_示范.pptx"

for d in (SVG, IMG):
    os.makedirs(d, exist_ok=True)
    for f in os.listdir(d):
        os.remove(os.path.join(d, f))

# ---------- 1) 旧 6 页 ----------
old_svg = os.path.join(OLD_WS, "svg_output")
old_img = os.path.join(OLD_WS, "images")
pages = sorted(f for f in os.listdir(old_svg) if f.endswith(".svg"))
for i, f in enumerate(pages, 1):
    shutil.copyfile(os.path.join(old_svg, f), os.path.join(SVG, "slide_%02d.svg" % i))
for f in os.listdir(old_img):
    shutil.copyfile(os.path.join(old_img, f), os.path.join(IMG, f))
print("旧页拷入", len(pages), "页，图片", len(os.listdir(IMG)), "张")

# ---------- 2) 新 4 页 ----------
SRC = [os.path.join(r"<ABS_PATH>", "img%02d.png" % i) for i in range(1, 8)]
SRC += [r"<ABS_PATH>",
        r"<ABS_PATH>"]
SRC = [p for p in SRC if os.path.exists(p)]

BG, INK, ACC, GREY = "#F6F4EE", "#1F2A2E", "#C3AE53", "#8A8578"
W, H = 1280, 720


def img(src_i, w, h, kind, name):
    shaped_image(SRC[src_i % len(SRC)], os.path.join(IMG, name), w, h, kind)
    return name


def text(x, y, s, size=22, color=INK, bold=False, anchor="start"):
    return ('<text x="%g" y="%g" font-size="%g" fill="%s" font-family="Microsoft YaHei"'
            '%s text-anchor="%s">%s</text>' % (
                x, y, size, color, ' font-weight="bold"' if bold else '', anchor, s))


def title(s, sub):
    return text(60, 64, s, 30, INK, True) + text(60, 96, sub, 16, GREY)


def page(parts):
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d" lang="zh-CN">\n'
            '<rect x="0" y="0" width="%d" height="%d" fill="%s"/>\n%s\n</svg>' % (W, H, W, H, W, H, BG, "\n".join(parts)))


new_pages = []

# P7 五种新形状
parts = [title("五种新形状（同尺寸 200×200）", "菱形·胶囊·波浪边·票券·破格矩形")]
kinds = [("rhombus", "rhombus 菱形"), ("capsule", "capsule 胶囊"), ("wave", "wave 波浪边"),
         ("ticket", "ticket 票券"), ("broken", "broken 破格矩形")]
x = 70
for i, (k, label) in enumerate(kinds):
    n = img(i, 200, 200, k, "n_%s.png" % k)
    parts.append('<image href="../images/%s" x="%d" y="190" width="200" height="200"/>' % (n, x))
    parts.append(outline_svg(x, 190, 200, 200, k))
    parts.append(text(x, 430, label, 18, GREY))
    x += 232
new_pages.append(page(parts))

# P8 三形咬合
parts = [title("三形咬合", "圆 + 六边 + 切角，压边 40px，后叠加白描边")]
a = img(0, 380, 380, "circle", "n_a.png"); b = img(1, 340, 340, "hex", "n_b.png"); c = img(2, 300, 300, "cut", "n_c.png")
parts.append('<image href="../images/%s" x="120" y="200" width="380" height="380"/>' % a)
parts.append(outline_svg(120, 200, 380, 380, "circle"))
parts.append('<image href="../images/%s" x="380" y="240" width="340" height="340" stroke="#FFFFFF" stroke-width="8"/>' % b)
parts.append(outline_svg(380, 240, 340, 340, "hex"))
parts.append('<image href="../images/%s" x="620" y="300" width="300" height="300" stroke="#FFFFFF" stroke-width="8"/>' % c)
parts.append(outline_svg(620, 300, 300, 300, "cut"))
new_pages.append(page(parts))

# P9 网格 + 破格
parts = [title("网格 + 破格", "3×3 切角网格，中心菱形放大溢出格线")]
cell, gap, x0, y0 = 160, 24, 200, 180
for r in range(3):
    for col in range(3):
        idx = r * 3 + col
        if idx == 4:
            continue
        n = img(idx, cell, cell, "cut", "n_g%d.png" % idx)
        px, py = x0 + col * (cell + gap), y0 + r * (cell + gap)
        parts.append('<image href="../images/%s" x="%d" y="%d" width="%d" height="%d"/>' % (n, px, py, cell, cell))
        parts.append(outline_svg(px, py, cell, cell, "cut", ACC, 2))
big = 240
bx = x0 + (cell + gap) + cell // 2 - big // 2
by = y0 + (cell + gap) + cell // 2 - big // 2
n = img(8, big, big, "rhombus", "n_big.png")
parts.append('<image href="../images/%s" x="%d" y="%d" width="%d" height="%d" stroke="#FFFFFF" stroke-width="10"/>' % (n, bx, by, big, big))
parts.append(outline_svg(bx, by, big, big, "rhombus"))
new_pages.append(page(parts))

# P10 同图多形状
parts = [title("同图多形状并置对比", "同一张图裁成四种形状并排，选型最直观")]
x = 90
for k, label in [("cut", "切角"), ("circle", "圆形"), ("hex", "六边"), ("arch", "拱形")]:
    n = img(3, 220, 220, k, "n_p4_%s.png" % k)
    parts.append('<image href="../images/%s" x="%d" y="210" width="220" height="220"/>' % (n, x))
    parts.append(outline_svg(x, 210, 220, 220, k))
    parts.append(text(x + 110, 470, label, 20, GREY, anchor="middle"))
    x += 280
new_pages.append(page(parts))

for i, s in enumerate(new_pages, len(pages) + 1):
    open(os.path.join(SVG, "slide_%02d.svg" % i), "w", encoding="utf-8").write(s)
print("新页写入", len(new_pages), "页，总计", len(pages) + len(new_pages))

subprocess.run([PY, os.path.join(SK, "scripts", "svg_quality_checker.py"), WS,
                "--quick-generate", "--canonical-authoring", "--stage", "final", "--json"], capture_output=True)
r = subprocess.run([PY, os.path.join(SK, "scripts", "svg_to_pptx.py"), WS, "-f", "ppt169",
                    "--quick-generate", "-t", "none", "-o", OUT],
                   capture_output=True, text=True, encoding="utf-8", errors="ignore")
print("export rc", r.returncode)
if r.returncode != 0:
    print((r.stdout or "")[-500:]); print((r.stderr or "")[-500:])
else:
    print("saved", OUT, round(os.path.getsize(OUT) / 1024), "KB")
