# -*- coding: utf-8 -*-
"""边框 → 原生形状映射（替代烘焙 PNG）

原则：
  图片不再是"裁好形状再贴上去"，而是 **形状 + 图片填充**：
      p:sp  +  a:prstGeom / a:custGeom  +  a:blipFill(原图)
  这样在 PowerPoint 里可以：改圆角/换图/改描边/编辑顶点。

用法：
    from frame_spec import FRAMES, sp_xml
    xml = sp_xml(shape_id=12, name="PicFrame", x=920, y=240, w=640, h=420,
                 frame="cut", rid="rId7", line="C3AE53", line_pt=2.25)
"""

# prst：PowerPoint 预设形状名；adj：预设形状可调值（0–100000）
# cust：需要自定义路径（由调用方按 custGeom 生成）
FRAMES = {
    # 基本型
    "cut":      dict(kind="prst", prst="snip2SameRect", adj={"adj1": 12000, "adj2": 0}),
    "arch":     dict(kind="prst", prst="round2SameRect", adj={"adj1": 40000, "adj2": 0}),
    "hex":      dict(kind="prst", prst="hexagon", adj={"adj": 25000, "vf": 115470}),
    "circle":   dict(kind="prst", prst="ellipse", adj={}),
    "rhombus":  dict(kind="prst", prst="diamond", adj={}),
    "capsule":  dict(kind="prst", prst="roundRect", adj={"adj": 50000}),
    "rect":     dict(kind="prst", prst="rect", adj={}),
    # 需要自定义路径（custGeom）
    "diag":     dict(kind="cust", note="斜切：矩形去掉左上/右下三角"),
    "arch_bar": dict(kind="cust", note="拱形 + 底部横条（并集）"),
    "duo_circle": dict(kind="cust", note="双圆并集（交点断开，只走外侧弧）"),
    "hex_bar":  dict(kind="cust", note="六边 + 左侧竖条（并集）"),
    "ring":     dict(kind="cust", note="外矩形剪除内矩形（镂空）"),
}

SP_TMPL = (
    '<p:sp xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}">'
    '<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    '<p:spPr>'
    '<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '{geom}'
    '<a:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
    '<a:ln w="{lw}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    '</p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody>'
    '</p:sp>'
)

NS = dict(p='http://schemas.openxmlformats.org/presentationml/2006/main',
          a='http://schemas.openxmlformats.org/drawingml/2006/main',
          r='http://schemas.openxmlformats.org/officeDocument/2006/relationships')


def prst_geom(prst, adj):
    gds = "".join('<a:gd name="%s" fmla="val %d"/>' % (k, v) for k, v in adj.items())
    return ('<a:prstGeom prst="%s"><a:avLst>%s</a:avLst></a:prstGeom>' % (prst, gds))


def sp_xml(shape_id, name, x, y, w, h, frame, rid, line="C3AE53", line_pt=2.25, geom_xml=None):
    """生成「形状 + 图片填充」的 p:sp；坐标单位 px（内部转 EMU）"""
    EMU = 9525
    spec = FRAMES[frame]
    if spec["kind"] == "prst":
        geom = prst_geom(spec["prst"], spec["adj"])
    else:
        if geom_xml is None:
            raise ValueError("边框 %s 需要自定义路径，请传入 geom_xml" % frame)
        geom = geom_xml
    return SP_TMPL.format(p=NS['p'], a=NS['a'], r=NS['r'],
                          sid=shape_id, name=name,
                          x=int(x * EMU), y=int(y * EMU), cx=int(w * EMU), cy=int(h * EMU),
                          geom=geom, rid=rid,
                          lw=int(line_pt * 12700), line=line)


if __name__ == "__main__":
    for k, v in FRAMES.items():
        print("%-11s %-6s %s" % (k, v["kind"], v.get("prst", v.get("note", ""))))

# ---------- custGeom 构造器（2026-11-14 补齐 diag / arch_bar / hex_bar / ring） ----------
def _pt(x, y):
    """百分比坐标（0–100000），越界自动夹紧"""
    return '<a:pt x="%d" y="%d"/>' % (max(0, min(100000, int(x))), max(0, min(100000, int(y))))


def _path(cmds):
    return ('<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            '<a:rect l="0" t="0" r="100000" b="100000"/>'
            '<a:pathLst><a:path w="100000" h="100000" fill="norm">%s</a:path></a:pathLst>'
            '</a:custGeom>' % "".join(cmds))


def geom_diag(cut=12000):
    """斜切：矩形去掉左上、右下两个三角"""
    c = cut
    return _path([
        '<a:moveTo>%s</a:moveTo>' % _pt(c, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 100000 - c),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000 - c, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(0, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(0, c),
        '<a:close/>',
    ])


def _arch_cmds(w, h, r):
    """拱形（上圆角、下直角）单条路径，坐标 0–1 比例"""
    X = lambda v: int(v * 100000)
    return [
        '<a:moveTo>%s</a:moveTo>' % _pt(X(0), X(1)),
        '<a:lnTo>%s</a:lnTo>' % _pt(X(0), X(r)),
        '<a:cubicBezTo>%s%s%s</a:cubicBezTo>' % (
            _pt(X(0), X(r * 0.45)), _pt(X(0.45 * w), X(0)), _pt(X(w / 2), X(0))),
        '<a:cubicBezTo>%s%s%s</a:cubicBezTo>' % (
            _pt(X(w - 0.45 * w), X(0)), _pt(X(w), X(r * 0.45)), _pt(X(w), X(r))),
        '<a:lnTo>%s</a:lnTo>' % _pt(X(w), X(1)),
        '<a:close/>',
    ]


def geom_arch_bar(bar=0.16):
    """拱形 + 底部横条（并集）：拱占上 1-bar，条在底部"""
    r = 0.30
    cmds = _arch_cmds(1.0, 1.0 - bar, r)
    # 底部横条（顺时针）
    cmds += [
        '<a:moveTo>%s</a:moveTo>' % _pt(0, int((1 - bar) * 100000)),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, int((1 - bar) * 100000)),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(0, 100000),
        '<a:close/>',
    ]
    return _path(cmds)


def geom_hex_bar(bar=0.16):
    """六边 + 左侧竖条（并集）"""
    b = int(bar * 100000)
    cmds = [
        '<a:moveTo>%s</a:moveTo>' % _pt(0, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(b, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(b, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(0, 100000),
        '<a:close/>',
    ]
    left = b
    mid = int(left + (100000 - left) * 0.25)
    cmds += [
        '<a:moveTo>%s</a:moveTo>' % _pt(mid, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000 - mid + left, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 50000),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000 - mid + left, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(mid, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(left, 50000),
        '<a:close/>',
    ]
    return _path(cmds)


def geom_ring(inner=0.16):
    """镂空：外矩形 − 内矩形（内框反向，形成洞）"""
    i = int(inner * 100000)
    j = 100000 - i
    return _path([
        '<a:moveTo>%s</a:moveTo>' % _pt(0, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 0),
        '<a:lnTo>%s</a:lnTo>' % _pt(100000, 100000),
        '<a:lnTo>%s</a:lnTo>' % _pt(0, 100000),
        '<a:close/>',
        '<a:moveTo>%s</a:moveTo>' % _pt(i, i),
        '<a:lnTo>%s</a:lnTo>' % _pt(i, j),
        '<a:lnTo>%s</a:lnTo>' % _pt(j, j),
        '<a:lnTo>%s</a:lnTo>' % _pt(j, i),
        '<a:close/>',
    ])


GEOM_BUILDERS = {
    "diag": geom_diag,
    "arch_bar": geom_arch_bar,
    "hex_bar": geom_hex_bar,
    "ring": geom_ring,
}
