from pptx import Presentation
from copy import copy

SRC = r'<USER_HOME>\Desktop\模板\大学万能职业生涯.pptx'
OUT = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\职业规划-大学版.pptx'

prs = Presentation(SRC)


def rewrite(tf, lines, white=False):
    """整段重写，字体字号沿用原来的第一个 run。"""
    p0 = tf.paragraphs[0]
    src_run = next((r for r in p0.runs if r.text.strip()), None)
    if src_run is None:
        return False
    font = src_run.font
    name, size, bold = font.name, font.size, font.bold
    color = None
    try:
        color = copy(font.color.rgb)
    except Exception:
        pass

    tf.clear()
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        if name:
            r.font.name = name
        if size:
            r.font.size = size
        if bold is not None:
            r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        if white:
            from pptx.dml.color import RGBColor
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return True


REPL = {
    3: [('身为当代的一名大学生', [
        '职业生涯规划不是一次决定，而是一套持续校准的方法。',
        '从自我认知出发，经过环境分析与评估调整，落到四年行动计划。',
        '目标很朴素：毕业时有一份对口实习、一个能讲清楚的作品。',
    ])],
    22: [('我将定期评估自己的职业目标', ['每学期末复盘：目标是否仍指向同一方向。'])],
    24: [('我将定期评估自己制定的实施策略', ['检查方法是否有效：作品、实习、表达。'])],
    26: [('我将根据自己的职业发展情况', ['路径可调整：连续两次受阻就换相近的路。'])],
    28: [('我还将评估其他因素', ['外部因素每年重评：行业周期、家庭、身体。'])],
    29: [('我还将评估其他因素', ['结论只写三行：守住、放弃、下一步。'])],
}

n = 0
for idx, rules in REPL.items():
    slide = prs.slides[idx - 1]
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        txt = sh.text_frame.text
        for old, lines in rules:
            if old in txt:
                if rewrite(sh.text_frame, lines, white=(idx == 3)):
                    n += 1
                break

prs.save(OUT)
print('rewrote', n, '->', OUT)
