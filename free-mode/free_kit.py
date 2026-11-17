# -*- coding: utf-8 -*-
"""free-mode 工具箱：开工作区、取参考件、改色、取景形状、填充、跨文件拼接、放图。

用法：
    python free_kit.py new     <主题>
    python free_kit.py add     <工作区> <效果编号>
    python free_kit.py recolor <工作区> <文件.pptx> <配色名> [--force]
    python free_kit.py shape   <图片> <形状> [宽 高]
    python free_kit.py fill    <工作区> <模板.pptx> <内容.json>
    python free_kit.py merge   <工作区> <目标.pptx> [源.pptx[:页,页] ...]
    python free_kit.py place   <工作区> <目标.pptx> <页码> <图片> [x y w [h]]   # cm
    python free_kit.py list

规则（与 free-mode/README.md 一致）：
- 只复制、不修改库内原件
- 产物一律落在工作区 out/，库内不留草稿
- merge / place 用 PowerPoint COM（跨文件合并属已批准的例外，不走 ppt-master）
"""
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import date
from pathlib import Path

# 库根：按本文件位置推导（free-mode/ 的上一级），换机器/换目录都不用改
LIB = Path(__file__).resolve().parents[1]
# 工作区：可用环境变量 FREE_WS 覆盖；默认放在 DSH 的 default-workspace 下
WS = Path(os.environ.get('FREE_WS') or
          (Path.home() / 'Documents' / '<DSH_HOME>' / 'default-workspace' / 'free'))
PALETTES = json.loads((LIB / 'palette-tools' / 'palettes.json').read_text('utf-8'))['palettes']


# ---------- 基础 ----------

def cmd_new(theme):
    d = WS / ('%s-%s' % (date.today().isoformat(), theme))
    for sub in ('draft', 'assets', 'out'):
        (d / sub).mkdir(parents=True, exist_ok=True)
    notes = d / 'notes.md'
    if not notes.exists():
        notes.write_text('# %s\n\n## 用了哪些源\n- \n\n## 改了什么\n- \n\n## 待办\n- \n' % theme,
                         encoding='utf-8')
    print('工作区:', d)
    return d


def cmd_add(ws, eff):
    ws = Path(ws)
    hits = [p for p in (LIB / 'effects').iterdir() if p.is_dir() and p.name.startswith(eff)]
    if not hits:
        sys.exit('找不到效果: %s' % eff)
    mod = hits[0]
    refs = sorted((mod / 'reference').glob('*.pptx'))
    if not refs:
        sys.exit('%s 没有参考件' % mod.name)
    dst = ws / 'assets'
    dst.mkdir(parents=True, exist_ok=True)
    for f in refs:
        shutil.copy2(f, dst / f.name)
        print('  复制 ->', f.name)
    print('来自', mod.name, '共', len(refs), '份')



# ---------- 预览与文字 ----------

def _sheet(pngs, out):
    from PIL import Image, ImageDraw
    ims = [Image.open(x).convert('RGB') for x in pngs if Path(x).exists()]
    if not ims:
        return None
    cols = min(3, len(ims))
    rows = (len(ims) + cols - 1) // cols
    cw, ch = 640, 360
    s = Image.new('RGB', (cols * cw, rows * (ch + 18)), (20, 20, 20))
    dr = ImageDraw.Draw(s)
    for i, im in enumerate(ims):
        x, y = (i % cols) * cw, (i // cols) * (ch + 18)
        s.paste(im.resize((cw, ch)), (x, y))
        dr.text((x + 6, y + ch + 3), 'p%d' % (i + 1), fill=(210, 210, 210))
    s.save(out)
    return out


def cmd_render(pptx, pages=None, outdir=None):
    """导出指定页为 PNG，并拼一张对照图。"""
    src = Path(pptx)
    if not src.exists():
        sys.exit('文件不存在: %s' % src)
    from pptx import Presentation
    total = len(Presentation(str(src)).slides)
    want = [int(x) for x in pages.split(',')] if pages else list(range(1, min(total, 9) + 1))
    want = [i for i in want if 1 <= i <= total]
    out = Path(outdir) if outdir else src.parent / (src.stem + '-preview')
    out.mkdir(parents=True, exist_ok=True)
    pngs = [out / ('p%02d.png' % i) for i in want]
    script = ("$ErrorActionPreference='Stop';"
              "$a=New-Object -ComObject PowerPoint.Application;"
              "$p=$a.Presentations.Open('%s',-1,0,0);" % src
              + ''.join("$p.Slides.Item(%d).Export('%s','PNG',1280,720);" % (i, str(out / ('p%02d.png' % i)))
                        for i in want)
              + "$p.Close();$a.Quit();Write-Output ('导出 %d 页')" % len(want))
    _ps(script)
    sheet = _sheet(pngs, out / 'sheet.png')
    print('预览目录:', out)
    if sheet:
        print('对照图:', sheet)
    return out



# ---------- 文字适配 ----------

FONT_CJK = r'<ABS_PATH>'


def _measure(text, pt, bold=False):
    """返回文本在 pt 字号下的宽度（单位≈pt）。"""
    from PIL import ImageFont
    try:
        f = ImageFont.truetype(FONT_CJK, max(1, int(round(pt))))
    except Exception:
        return len(text) * pt
    w = f.getlength(text)
    return w * (1.02 if bold else 1.0)


def _is_vertical(shape):
    try:
        from lxml import etree
        body = shape.text_frame._txBody.find(
            '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
        return body is not None and body.get('vert') in ('eaVert', 'vert', 'vert270')
    except Exception:
        return False


def fit_shapes(pptx, only=None, min_pt=9.0, floor_ratio=0.30):
    """对每段文字做适配：放不下就整体缩字号（不低于 min_pt 与 floor_ratio 倍）。"""
    from pptx import Presentation
    from pptx.util import Pt
    prs = Presentation(str(pptx))
    adjusted = []

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
            else:
                yield sh

    EMU_PT = 12700.0
    for si, slide in enumerate(prs.slides, 1):
        for sh in walk(slide.shapes):
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            if not tf.text.strip():
                continue
            if only is not None and (si, sh.shape_id) not in only:
                continue
            vert = _is_vertical(sh)
            avail = ((sh.height if vert else sh.width) / EMU_PT
                     - (tf.margin_top + tf.margin_bottom if vert else tf.margin_left + tf.margin_right) / EMU_PT)
            if avail <= 6:
                continue
            for para in tf.paragraphs:
                runs = [r for r in para.runs if (r.text or '').strip()]
                if not runs:
                    continue
                sizes = [r.font.size.pt if r.font.size else 18.0 for r in runs]
                need = sum(_measure(r.text, s, bool(r.font.bold)) for r, s in zip(runs, sizes))
                if need <= avail * 0.98:
                    continue
                base = max(sizes)
                ratio = max(floor_ratio, avail * 0.96 / need)
                new = max(min_pt, base * ratio)
                for r, s in zip(runs, sizes):
                    r.font.size = Pt(round(s * new / base, 1))
                adjusted.append((si, sh.name[:14], round(base, 1), round(new, 1)))
    if adjusted:
        prs.save(str(pptx))
    return adjusted


def cmd_text(pptx, *pairs):
    """批量替换文字：old=new，精确匹配整段 run，保留格式。"""
    from pptx import Presentation
    src = Path(pptx)
    if not src.exists():
        sys.exit('文件不存在: %s' % src)
    rules = []
    for pr in pairs:
        if '=' not in pr:
            sys.exit('格式应为 old=new：%s' % pr)
        a, b = pr.split('=', 1)
        rules.append((a, b))
    prs = Presentation(str(src))
    hits = {a: 0 for a, _ in rules}

    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
            else:
                yield sh

    touched = set()
    for si, slide in enumerate(prs.slides, 1):
        for sh in walk(slide.shapes):
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    for a, b in rules:
                        if r.text == a:
                            r.text = b
                            hits[a] += 1
                            touched.add((si, sh.shape_id))
                # 整段匹配（这类模板常把“标题1”拆成多个 run）
                whole = ''.join((r.text or '') for r in para.runs)
                for a, b in rules:
                    if whole == a and para.runs:
                        para.runs[0].text = b
                        for r2 in para.runs[1:]:
                            r2.text = ''
                        hits[a] += 1
                        touched.add((si, sh.shape_id))
    out = src.parent / (src.stem + '-text.pptx')
    prs.save(str(out))

    if '--no-fit' not in sys.argv:
        adj = fit_shapes(out, only=touched)
        if adj:
            print('  自动缩字号 %d 处（放不下，按比例缩小，不换行）：' % len(adj))
            for si, name, a, b in adj[:6]:
                print('    第%d页 %s：%spt -> %spt' % (si, name, a, b))
        else:
            print('  文字都放得下，未调整字号')
    for a, b in rules:
        print('  %r -> %r  %d 处' % (a, b, hits[a]))
        if len(b) > len(a) + 2 and hits[a]:
            print('    ⚠ 新文本比原文长 %d 字，目标框可能放不下（会换行/竖排），记得看预览' % (len(b) - len(a)))
    print('产物 ->', out)
    return out


def cmd_list():
    print('可用效果源:')
    for d in sorted((LIB / 'effects').iterdir()):
        refs = d / 'reference'
        if refs.is_dir() and list(refs.glob('*.pptx')):
            print('  ', d.name)
    print('可用配色:')
    for k, v in PALETTES.items():
        print('   %-16s %s' % (k, v['name']))


# ---------- 改色 ----------

def dominant_colors(pptx, top=6):
    z = zipfile.ZipFile(pptx)
    c = Counter()
    for n in z.namelist():
        if re.match(r'ppt/slides/slide\d+\.xml$', n):
            c.update(re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', z.read(n).decode('utf-8', 'ignore')))
    return [k for k, _ in c.most_common(40)
            if k.upper() not in ('FFFFFF', '000000', 'F2F2F2', 'FAFAFA')][:top]


def tint(hexc, amount):
    r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
    return '%02X%02X%02X' % (int(r + (255 - r) * amount),
                             int(g + (255 - g) * amount),
                             int(b + (255 - b) * amount))


def shade(hexc, amount):
    r, g, b = (int(hexc[i:i + 2], 16) for i in (0, 2, 4))
    return '%02X%02X%02X' % (int(r * (1 - amount)), int(g * (1 - amount)), int(b * (1 - amount)))


def color_source_of(filename):
    """按参考件文件名反查效果 spec 的 color_source。"""
    for spec in LIB.glob('effects/*/spec.json'):
        refs = spec.parent / 'reference'
        if refs.is_dir() and (refs / filename).exists():
            try:
                return json.loads(spec.read_text('utf-8')).get('color_source', '')
            except Exception:
                return ''
    return ''


def cmd_recolor(ws, name, palette):
    ws = Path(ws)
    src = ws / 'assets' / name
    if not src.exists():
        cand = list((ws / 'assets').glob(name))
        if not cand:
            sys.exit('assets 里没有: %s' % name)
        src = cand[0]

    force = '--force' in sys.argv
    cs = color_source_of(src.name)
    if cs == 'image' and not force:
        print('这个效果的主体颜色来自图片（color_source=image），改色基本无效：%s' % src.name)
        print('建议：换图、换效果，或加 --force 强行跑一遍。')
        sys.exit(2)

    if palette not in PALETTES:
        by_name = [k for k, v in PALETTES.items() if v['name'] == palette]
        if not by_name:
            sys.exit('没有这个配色: %s（用 list 查看）' % palette)
        palette = by_name[0]
    p = PALETTES[palette]
    primary, dark, accent = p['primary'], p['dark'], p['accent']

    def hsv(hexc):
        r, g, b = (int(hexc[i:i + 2], 16) / 255 for i in (0, 2, 4))
        mx, mn = max(r, g, b), min(r, g, b)
        return (0 if mx == 0 else (mx - mn) / mx), mx

    mapping = {}
    for c in dominant_colors(src):
        s, v = hsv(c)
        if v > 0.90 and s < 0.20:
            mapping[c] = tint(primary, 0.92)
        elif v < 0.25:
            mapping[c] = dark
        elif s > 0.45 and v > 0.80:
            mapping[c] = accent
        elif s > 0.45:
            mapping[c] = primary
        elif v > 0.62:
            mapping[c] = tint(primary, 0.72)
        elif v > 0.38:
            mapping[c] = tint(primary, 0.40)
        else:
            mapping[c] = primary

    out = (ws / 'out' / ('%s-%s' % (src.stem, p['name']))).with_suffix('.pptx')
    shutil.copyfile(src, out)
    with zipfile.ZipFile(out) as z:
        items = {n: z.read(n) for n in z.namelist()}

    n = 0
    for k in list(items):
        if not re.match(r'ppt/(slides/slide\d+\.xml|slideMasters/.+|slideLayouts/.+|theme/.+)', k):
            continue
        s = items[k].decode('utf-8', 'ignore')
        o = s
        for old, new in mapping.items():
            s = re.sub(r'(srgbClr val=")' + old + r'(")', r'\g<1>' + new + r'\g<2>', s, flags=re.I)
        if s != o:
            items[k] = s.encode('utf-8')
            n += 1

    slots = {'accent1': primary, 'accent2': accent, 'accent3': tint(primary, 0.60),
             'accent4': tint(primary, 0.80), 'accent5': shade(primary, 0.20), 'accent6': dark,
             'dk2': dark, 'lt2': tint(primary, 0.85)}
    tn = 0
    for k in list(items):
        if not re.match(r'ppt/theme/theme\d+\.xml$', k):
            continue
        s = items[k].decode('utf-8', 'ignore')
        for slot, col in slots.items():
            s = re.sub(r'(<a:%s><a:srgbClr val=")[0-9A-Fa-f]{6}(")' % slot,
                       r'\g<1>' + col + r'\g<2>', s)
        items[k] = s.encode('utf-8')
        tn += 1

    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in items.items():
            zi = zipfile.ZipInfo(k, date_time=(2020, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, v)
    print('配色 %s：%d 个部件 + %d 个主题改写 -> %s' % (p['name'], n, tn, out))
    if mapping:
        print('映射:', ', '.join('%s→%s' % (a, b) for a, b in mapping.items()))
    return out


# ---------- 取景形状 ----------

def cmd_shape(img, kind, w=900, h=600):
    sys.path.insert(0, str(LIB / 'image-tools' / 'tools'))
    try:
        from mask_lib import shaped_image
    except Exception as e:
        sys.exit('无法导入 mask_lib：%s' % e)
    out = Path.cwd() / ('%s-%s.png' % (Path(img).stem, kind))
    shaped_image(str(img), str(out), int(w), int(h), kind)
    print('形状图 ->', out)
    return out


# ---------- PowerPoint COM ----------

def _ps(script):
    r = subprocess.run(['powershell', '-NoProfile', '-Command', script],
                       capture_output=True, text=True, encoding='utf-8', errors='ignore')
    out = (r.stdout or '').strip()
    if out:
        print(out[-600:])
    if r.returncode != 0:
        print((r.stderr or '').strip()[-500:])
        sys.exit(r.returncode)
    return out


def cmd_fill(ws, pptx, content_json):
    """按占位名批量填文字/图片（复用 template-tools/fill_template.py）。"""
    ws = Path(ws)
    tpl = Path(pptx) if Path(pptx).is_absolute() else ws / 'assets' / pptx
    if not tpl.exists():
        sys.exit('模板不存在: %s' % tpl)
    script = LIB / 'template-tools' / 'tools' / 'fill_template.py'
    if not script.exists():
        sys.exit('找不到 fill_template.py')
    out = ws / 'out' / (tpl.stem + '-filled.pptx')
    out.parent.mkdir(parents=True, exist_ok=True)

    cj = Path(content_json)
    if not cj.is_absolute():
        cj = ws / content_json
    norm = ws / 'draft' / (cj.stem + '-nobom.json')
    norm.parent.mkdir(parents=True, exist_ok=True)
    norm.write_text(cj.read_text('utf-8-sig'), encoding='utf-8')

    r = subprocess.run([sys.executable, '-X', 'utf8', str(script), str(tpl), str(norm), str(out)],
                       capture_output=True, text=True, encoding='utf-8', errors='ignore')
    print((r.stdout or '').strip()[-500:])
    if r.returncode != 0:
        print((r.stderr or '').strip()[-400:])
        sys.exit(r.returncode)
    print('填充产物 ->', out)
    return out


def _slide_count(pptx):
    from pptx import Presentation
    return len(Presentation(str(pptx)).slides)



# ---------- 接缝检查 ----------

def _slide_names(pptx, index):
    """取第 index 页（1-based）的所有形状名。"""
    from pptx import Presentation
    prs = Presentation(str(pptx))
    slide = prs.slides[index - 1]
    names = set()

    def walk(shapes):
        for sh in shapes:
            names.add(sh.name)
            if sh.shape_type == 6:          # group
                walk(sh.shapes)

    walk(slide.shapes)
    return names


def _has_morph(pptx, index):
    """第 index 页（1-based）切换是否为 morph。"""
    import re as _re
    import zipfile
    from pptx import Presentation
    prs = Presentation(str(pptx))
    part = prs.slides[index - 1].part.partname
    with zipfile.ZipFile(pptx) as z:
        x = z.read(str(part).lstrip('/')).decode('utf-8', 'ignore')
    tr = _re.search(r'<p:transition[^>]*>.*?</p:transition>', x, _re.S)
    return bool(tr and 'morph' in tr.group(0))


def inspect_seams(pptx, seam_pages, fix=False):
    """检查每个接缝页与上一页的同名对象数；可选把接缝 morph 改成淡入。"""
    lines, fixed = [], []
    for idx in seam_pages:
        if idx <= 1:
            continue
        prev_names = _slide_names(pptx, idx - 1)
        cur_names = _slide_names(pptx, idx)
        common = len(prev_names & cur_names)
        morph = _has_morph(pptx, idx)
        if not morph:
            lines.append('第%d页：切换不是 morph，接缝无风险（同名对象 %d 个）' % (idx, common))
            continue
        if common >= 2:
            lines.append('第%d页：morph + 同名对象 %d 个，动效应正常' % (idx, common))
        else:
            lines.append('第%d页：morph 但同名对象仅 %d 个 → 会退化成淡入' % (idx, common))
            if fix:
                to_fade(pptx, idx)
                fixed.append(idx)
    return lines, fixed


def to_fade(pptx, index):
    """把第 index 页的切换改成淡入（0.7 秒）。"""
    import re as _re
    import zipfile
    from pptx import Presentation
    prs = Presentation(str(pptx))
    part = str(prs.slides[index - 1].part.partname).lstrip('/')
    with zipfile.ZipFile(pptx) as z:
        items = {n: z.read(n) for n in z.namelist()}
    x = items[part].decode('utf-8', 'ignore')
    new_tr = ('<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
              'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
              'spd="med"><p:fade/></p:transition>')
    x2, n = _re.subn(r'<p:transition[^>]*>.*?</p:transition>', new_tr, x, flags=_re.S)
    if n == 0:
        return False
    items[part] = x2.encode('utf-8')
    with zipfile.ZipFile(pptx, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in items.items():
            zi = zipfile.ZipInfo(k, date_time=(2020, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, v)
    return True


def cmd_merge(ws, target, *sources):
    """跨文件拼接：把 sources 的页并入 target（例外：PowerPoint COM，不走 ppt-master）。

    InsertFromFile(File, Index, SlideStart, SlideEnd)：Index 为 1-based 插入位置，
    SlideStart/SlideEnd 为源文件 1-based 页号。指定页时先整份插入再删多余页。
    """
    ws = Path(ws)
    sources = tuple(s for s in sources if not s.startswith('--'))
    tgt = Path(target) if Path(target).is_absolute() else ws / 'assets' / target
    if not tgt.exists():
        sys.exit('目标不存在: %s' % tgt)
    pairs = []
    for s in sources:
        name, _, pages = s.partition(':')
        src = Path(name) if Path(name).is_absolute() else ws / 'assets' / name
        if not src.exists():
            sys.exit('源不存在: %s' % src)
        keep = [int(x) for x in pages.split(',')] if pages else []
        pairs.append((src, keep, _slide_count(src)))

    out = ws / 'out' / (tgt.stem + '-merged.pptx')
    out.parent.mkdir(parents=True, exist_ok=True)
    work = ws / 'draft' / ('merge-' + tgt.name)
    shutil.copy2(tgt, work)

    parts = ["$ErrorActionPreference='Stop';",
             "$a=New-Object -ComObject PowerPoint.Application;",
             "$p=$a.Presentations.Open('%s',-1,0,0);" % work]
    for src, keep, total in pairs:
        parts.append("$before=$p.Slides.Count;")
        parts.append("$p.Slides.InsertFromFile('%s',$before,1,%d) | Out-Null;" % (src, total))
        if keep:
            allow = ','.join(str(x) for x in keep)
            parts.append("$after=$p.Slides.Count;")
            parts.append("for($i=$after;$i -gt $before;$i--){"
                         "$k=$i-$before;if(@(%s) -notcontains $k){$p.Slides.Item($i).Delete()}};" % allow)
    parts.append("$p.SaveAs('%s');" % out)
    parts.append("Write-Output ('页数=' + $p.Slides.Count);")
    parts.append("$p.Close();$a.Quit()")
    _ps(''.join(parts))
    print('拼接产物 ->', out)

    # ---- 接缝检查 ----
    seam_pages = []
    cursor = _slide_count(tgt)
    for src, keep, total in pairs:
        added = len(keep) if keep else total
        seam_pages.append(cursor + 1)          # 每个新块的第一页
        cursor += added

    fix = '--fix-seams' in sys.argv
    lines, fixed = inspect_seams(out, seam_pages, fix)
    print()
    print('--- 拼接后自检 ---')
    for ln in lines:
        print('   ' + ln)
    print('--- 提醒 ---')
    print('   1) 动画（p:timing）与切换都随页保留；但 morph 靠“相邻页同名对象”配对，')
    print('      接缝处若没有同名对象，会退化成整体淡入，原本的滑转/位移看不出来。')
    print('   2) 源文件的母版/版式会被一并带入，外观不变但文件变大。')
    print('   3) 原页里的“自动换片时间”也会保留，可能打乱节奏。')
    print('   4) 交付前请按 F5 从头放一遍，重点看每个接缝的前后两页。')
    if fixed:
        print('   已把 %s 的切换改成淡入（--fix-seams）。' % '、'.join('第%d页' % x for x in fixed))
    elif any('morph' in ln for ln in lines):
        print('   这次没有自动修；需要的话加 --fix-seams 重跑，接缝处会改成淡入。')
    return out


def cmd_place(ws, pptx, page, img, x=2.0, y=2.0, w=10.0, h=None):
    """把图片放进指定页（cm 坐标）。"""
    ws = Path(ws)
    src = Path(pptx)
    if not src.is_absolute():
        for base in (ws / 'assets', ws / 'out', Path.cwd()):
            cand = base / pptx
            if cand.exists():
                src = cand
                break
        else:
            src = ws / 'assets' / pptx
    pic = Path(img)
    if not src.exists():
        sys.exit('PPT 不存在: %s' % src)
    if not pic.exists():
        cand = list(Path.cwd().glob(img))
        if not cand:
            sys.exit('图片不存在: %s' % pic)
        pic = cand[0]
    out = ws / 'out' / (src.stem + '-placed.pptx')
    out.parent.mkdir(parents=True, exist_ok=True)
    work = ws / 'draft' / ('place-' + src.name)
    shutil.copy2(src, work)

    from PIL import Image as _Image
    iw, ih = _Image.open(pic).size
    x, y = float(x), float(y)
    w = float(w)
    h = float(h) if h else w * ih / iw
    CM = 28.3465
    script = ("$ErrorActionPreference='Stop';"
              "$a=New-Object -ComObject PowerPoint.Application;"
              "$p=$a.Presentations.Open('%s',-1,0,0);" % work
              + "$s=$p.Slides.Item(%d);" % int(page)
              + "$s.Shapes.AddPicture('%s',0,-1,%d,%d,%d,%d) | Out-Null;"
                % (pic, x * CM, y * CM, w * CM, h * CM)
              + "$p.SaveAs('%s');" % out
              + "Write-Output ('已放置到第 %d 页');" % int(page)
              + "$p.Close();$a.Quit()")
    _ps(script)
    print('产物 ->', out)
    return out


if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    cmd = sys.argv[1]
    if cmd == 'new':
        cmd_new(sys.argv[2])
    elif cmd == 'add':
        cmd_add(sys.argv[2], sys.argv[3])
    elif cmd == 'recolor':
        cmd_recolor(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == 'shape':
        cmd_shape(sys.argv[2], sys.argv[3], *sys.argv[4:6])
    elif cmd == 'fill':
        cmd_fill(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == 'merge':
        cmd_merge(sys.argv[2], sys.argv[3], *sys.argv[4:])
    elif cmd == 'place':
        cmd_place(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], *sys.argv[6:9])
    elif cmd == 'render':
        pages = None
        outdir = None
        rest = sys.argv[3:]
        if rest and not rest[0].startswith('-'):
            pages = rest[0]
        if '--out' in rest:
            outdir = rest[rest.index('--out') + 1]
        cmd_render(sys.argv[2], pages, outdir)
    elif cmd == 'text':
        cmd_text(sys.argv[2], *sys.argv[3:])
    elif cmd == 'list':
        cmd_list()
    else:
        sys.exit(__doc__)
