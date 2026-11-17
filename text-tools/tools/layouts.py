# -*- coding: utf-8 -*-
"""版式封装：时间轴 / 数据块 / 引文卡（一次调用拿到形状清单）

设计：
    每个函数返回 list[dict]，字段 x,y,w,h（px）、fill、line、text、size、color、bold、kind
    emit(spTree, shapes) 把它们写成 p:sp，追加到指定 spTree

用法：
    from layouts import timeline, datablocks, quote, emit
    shapes = timeline([("2026.11", "文献研读与门址考据"), ...], x=470, y0=248)
    emit(spTree, shapes)
"""
from lxml import etree

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 9525

GOLD = "C3AE53"
INK = "33404D"
TINT = "EAF1FA"
EDGE = "DCE8F6"
DARK = "17335F"


def _rect(x, y, w, h, fill, line=None, adj=6000, kind="round"):
    return dict(kind=kind, x=x, y=y, w=w, h=h, fill=fill, line=line, adj=adj)


def _text(x, y, w, h, text, size, color, bold=False, align="l"):
    return dict(kind="text", x=x, y=y, w=w, h=h, text=text, size=size,
                color=color, bold=bold, align=align)


def timeline(items, x=470, y0=248, step=92, line_color="D8DFC9",
             chip_w=430, chip_h=40, date_gap=180, dot=14,
             date_size=16, item_size=19):
    """竖向时间轴：线 + 圆点 + 胶囊 + 日期（线左侧右对齐）"""
    n = len(items)
    out = [_rect(x, y0 - 10, 3, step * (n - 1) + 30, line_color, kind="rect")]
    for i, (when, what) in enumerate(items):
        y = y0 + i * step
        out.append(_rect(x - dot // 2, y - dot // 2, dot, dot, GOLD, kind="ellipse"))
        out.append(_rect(x + 34, y - chip_h // 2, chip_w, chip_h, TINT, adj=50000))
        out.append(_text(x - date_gap, y - chip_h // 2, date_gap - 20, chip_h,
                         when, date_size, "8A8578", align="r"))
        out.append(_text(x + 52, y - chip_h // 2, chip_w - 26, chip_h, what, item_size, INK))
    return out


def datablocks(values, x=198, y=226, card_w=210, card_h=140, gap=24,
                total_text=None, total_w=678, total_h=74, total_gap=26,
                num_size=30, label_size=15, total_size=22):
    """数据块：n 张数字卡 + 可选合计条"""
    out = []
    for i, (num, label) in enumerate(values):
        cx = x + i * (card_w + gap)
        out.append(_rect(cx, y, card_w, card_h, "FFFFFF", line=EDGE))
        out.append(_text(cx, y + 22, card_w, 52, num, num_size, "1F4E79", bold=True, align="ctr"))
        out.append(_text(cx, y + 74, card_w, 32, label, label_size, "8A8578", align="ctr"))
    if total_text:
        ty = y + card_h + total_gap
        out.append(_rect(x, ty, total_w, total_h, TINT, line=EDGE))
        out.append(_text(x, ty, total_w, total_h, total_text, total_size, DARK, bold=True, align="ctr"))
    return out


def quote(text, body, x=198, y=210, qw=470, qh=132,
          card_w=700, card_h=110, card_gap=12,
          quote_size=17, body_size=18, source="—— 摘录"):
    """引文卡（小）+ 若干正文卡（大，带序号块）"""
    out = [_rect(x, y, qw, qh, TINT, line=EDGE)]
    out.append(_text(x + 16, y + 6, 40, 40, "“", 30, GOLD, bold=True))
    out.append(_text(x + 46, y + 16, qw - 60, 34, text, quote_size, "1F2A2E"))
    out.append(_text(x + 46, y + qh - 34, 200, 24, source, 12, "8A8578"))
    by = y + qh + 30
    for i, t in enumerate(body):
        cy = by + i * (card_h + card_gap)
        out.append(_rect(x, cy, card_w, card_h, "FFFFFF", line=EDGE))
        out.append(_rect(x + 16, cy + 22, 34, 34, "1E447A", line=GOLD, adj=20000))
        out.append(_text(x + 16, cy + 26, 34, 28, "0%d" % (i + 1), 14, "FFFFFF", bold=True, align="ctr"))
        out.append(_text(x + 64, cy + 26, card_w - 90, 60, t, body_size, INK))
    return out


def sp_xml(sid, s):
    x, y = int(s["x"] * EMU), int(s["y"] * EMU)
    w, h = int(s["w"] * EMU), int(s["h"] * EMU)
    if s["kind"] == "text":
        fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % s["color"])
        return ('<p:sp xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
                '<p:nvSpPr><p:cNvPr id="%d" name="T%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                '<p:txBody><a:bodyPr wrap="none" anchor="ctr"/><a:lstStyle/>'
                '<a:p><a:pPr algn="%s"/><a:r><a:rPr lang="zh-CN" sz="%d" b="%d">%s</a:rPr>'
                '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
                % (P, A, R, sid, sid, x, y, w, h,
                   {"l": "l", "r": "r", "ctr": "ctr"}[s.get("align", "l")],
                   int(s["size"] * 100), 1 if s.get("bold") else 0, fill, s["text"]))
    prst = "ellipse" if s["kind"] == "ellipse" else "roundRect" if s["kind"] == "round" else "rect"
    av = '<a:gd name="adj" fmla="val %d"/>' % int(s.get("adj", 6000)) if prst == "roundRect" else ''
    ln = ('<a:ln w="25400"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % s["line"]
          if s.get("line") else '<a:ln><a:noFill/></a:ln>')
    return ('<p:sp xmlns:p="%s" xmlns:a="%s" xmlns:r="%s">'
            '<p:nvSpPr><p:cNvPr id="%d" name="S%d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>'
            '</p:sp>'
            % (P, A, R, sid, sid, x, y, w, h, prst, av, s["fill"], ln))


def emit(spTree, shapes, base_id=2000):
    """把形状清单写入 spTree"""
    sid = base_id
    for s in shapes:
        spTree.append(etree.fromstring(sp_xml(sid, s)))
        sid += 1
    return len(shapes)


if __name__ == '__main__':
    tl = timeline([("2026.11", "文献研读"), ("2027.01", "实地踏勘")])
    db = datablocks([("425.45", "交通费"), ("750.00", "印刷物料")], total_text="合计 1175.45 元")
    q = quote("十门十声：每座城门各采一段", ["访谈录音加密备份", "隐私内容遮蔽处理"])
    print('时间轴', len(tl), '个形状；数据块', len(db), '；引文卡', len(q))
