import os, subprocess, glob, json
from PIL import Image
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

F = r'<USER_HOME>\<OUTPUT_DIR>\效果9.pptx'
NODE = r'<DSH_CHECKOUT>\resources\runtime\primary-runtime\dependencies\node\bin\node.exe'
CLI = r'<DSH_CHECKOUT>\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js'
TMP = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\<TMP>'

prs = Presentation(F)
n = len(prs.slides)
pages = ','.join(str(i) for i in range(1, n + 1))
subprocess.run([NODE, CLI, 'render', '--input', F, '--output-dir', TMP, '--pages', pages],
               capture_output=True)

bright = {}
for i in range(1, n + 1):
    p = os.path.join(TMP, 'page-%04d.png' % i)
    if not os.path.exists(p):
        bright[i] = 200
        continue
    im = Image.open(p).convert('L').resize((64, 36))
    px = list(im.getdata())
    bright[i] = sum(px) / len(px)

dark_pages = [i for i, b in bright.items() if b < 130]
print('dark pages', dark_pages)
print('brightness sample', {k: round(v) for k, v in list(bright.items())[:8]})

changed = 0
for idx, slide in enumerate(prs.slides, start=1):
    white = bright[idx] < 130
    color = RGBColor(0xFF, 0xFF, 0xFF) if white else RGBColor(0x00, 0x00, 0x00)

    def walk(shapes):
        global changed
        for sh in shapes:
            if sh.shape_type == 6:
                walk(sh.shapes)
                continue
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            if not tf.text.strip():
                continue
            for p in tf.paragraphs:
                for r in p.runs:
                    if not r.text.strip():
                        continue
                    size = r.font.size.pt if r.font.size else 18
                    label = '标题' if size >= 28 else '正文'
                    r.text = '%s：请输入文本' % label
                    r.font.color.rgb = color
                    changed += 1

    walk(slide.shapes)

prs.save(F)
print('runs changed', changed)
