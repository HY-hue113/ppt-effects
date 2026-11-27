# TOOLS · 两套工具的使用逻辑

> 库根索引文档。**持续完善**——新增形状/底卡/规则都回写这里。
> 详细参数见 `image-tools/README.md` 与 `text-tools/README.md`。

---

## 〇、palette-tools：选色

**偏学术类必问**：4 套主推（学术蓝 / 红 / 绿 / 紫）+ 7 套学科配色（橄榄黄绿 / 青蓝湖色 / 灰紫丁香 / 暖橙金棕 / 砖红暖棕 / 森系墨绿 / 深海蓝系），共 11 套（`palette-tools/README.md`）。

```python
from palette import get, apply_theme
theme = apply_theme("academic-blue")   # 语义色：primary/secondary/light/dark/accent/nav_bar_dark/card_edge/chip_bg
```

`tint()` / `shade()` 派生浅色、深色、分隔线、蒙版底色。

## 一、image-tools：图片取景

> **慎用（2026-11-15 起）**：动效模式与学术模式的参考件已内建取景形状，
> 直接改图片填充即可；**不要轻易启用本区**，只有用户明确要求新做特殊轮廓时才引入。


**核心原则：图片是形状的填充，不是裁好的 PNG。**

| 场景 | 做法 |
|---|---|
| 常规取景 | `frame_spec.sp_xml(..., frame="cut"/"arch"/"hex"/"circle"/"capsule", rid=…)` → `p:sp + prstGeom + blipFill` |
| 复杂轮廓 | 并集/相交/剪除/自由造型：PowerPoint 的「合并形状」「编辑顶点」，或脚本生成 `custGeom` 后过 `custgeom_check` |
| 旧路径 | `mask_lib.shaped_image()`（烘焙 PNG）**仅用于非 PPT 场景**，PPT 内不再使用 |

**调用顺序**：定照片 → 定轮廓 → 建独立 rels（不猜页面已有 rels）

**custGeom 三条硬要求**（`tools/custgeom_check.py` 校验）：
1. `<a:path fill="norm">`
2. `spPr` 顺序 `xfrm → geom → 填充 → ln`
3. 路径闭合、坐标 ∈ [0,100000]

**纪律**：同款边框单用 ≤2 次；宽画幅只用不变形形状（capsule/cut）；圆用 `ellipse` 不拉伸。

---

## 一·五、layout-tools：版式骨架（独立保留）

> 本区不属于动效/学术两种模式的流程，独立保留，库内流程不引用。

**22 页原生版式**（`layout-tools/reference/版式库-参考件-22页.pptx`），246 个形状 + 2 张可换色图片。

| 场景 | 做法 |
|---|---|
| 要一套汇报骨架 | 复制所需页面，图片位右键「设置形状格式 → 填充 → 图片」 |
| 换主色 | 形状是实色，直接改填充；两张图片用双色调（`tools/picture_duotone.py`） |
| 加页 | 复制同类版面页，不要重建形状 |

**踩坑记录**（详见 `layout-tools/README.md`）：栅格页要分「整页图」与「局部图」两类；局部图只换图片，不能整页清空；大图占位要沉底；弧线用拟合圆而非像素追踪。

---

## 二、text-tools：文本底卡

**核心原则：标题与正文分开承托，正文 ≥20pt。**

| 用途 | 容器 |
|---|---|
| 小标题 | `tab()` 书签式，独立成块 |
| 正文块 | `card()`：plain / tint / bar / line / **corner / duo / topbar**（方卡分型） |
| 单行要点 | `chip()`（带金点） |
| 序号 | `badge()` |

**组合规则**
- 标题与正文**不共用一张卡**
- 正文不一定要一张大卡：3–4 条 → 一条一卡；更多 → 2×2 网格
- 混排 = 强调卡（图 + tint）+ 胶囊 + 小标签

**字号**：≥20pt；装不下就减字数或加大容器，**不缩字号**；用 `<tspan>` 显式分行

**卡内填图**：必须加**局部蒙版**（尺寸=卡片、不外溢）；亮底白蒙版+深字，暗底深蒙版+白字

**纪律**：同款底卡不连续三页同形式（用分型轮换）

---

## 三、两者配合的固定 z 序

```
底层：卡片形状（可含图片填充）
中层：蒙版（仅当卡内有图）
上层：文字
旁侧：取景图（shape + blipFill）独立成块
```

---

## 四、通用规则（两套共用）

| 规则 | 说明 |
|---|---|
| 尺寸非定值 | 坐标只是当前取值；按「文案字数 → 字号下限 → 可用区域」推 |
| 形状优先预设 | 圆/圆角/切角用 `prstGeom`；手写路径必须过校验 |
| 图片独立注入 | 媒体 + rels 显式建立，不依赖页面原有图片 |
| 入库前给预览 | 用户确认后才 `git commit` |

---

## 四·五、template-tools：模板复刻

**成果**：`effects/academic/templates/学术模板1/reference/` 六套版式（各 23 页），详见 `template-tools/README.md`。

**用途**：把现成 PPT 模板转成**可编辑占位版**。

**六步流程**（`rebuild.py`）：

| 步 | 动作 |
|---|---|
| 1 | 删前 3 页前置说明（视情况，页数由用户确认） |
| 2 | 其余页保持原样 |
| 3 | 前 5 页与尾页只改标题 |
| 4 | 内容页文字 → 「标题 / 文本」 |
| 5 | 图片 → 「图片」框（宽 ≥1100 的背景图保留）；图表 → 「图表」框 |
| 6 | 校徽由用户手动在母版替换（视图 → 幻灯片母版） |

**保留原文的字段**：指导教师 / 答辩人 / 日期 / 课题小组 / 研究方向 等表单字段。

**坐标表**：`template-tools/logo_boxes.json`（每个模板的校徽/校名位置）。

## 五、库结构总览

```
ppt-effects/
├─ effects/                动效类 001–022
│   └─ academic/           偏学术类（学术导航0…）
├─ palette-tools/          选色（4 套主推 + 7 套学科 = 11 套）
├─ image-tools/            图片取景（预设几何 / custGeom / 校验）
├─ text-tools/             文本三分类：title / nav / textbox（190 种样式）
├─ layout-tools/           版式区（独立保留）
├─ template-tools/         模板复刻（占位化 + 校徽/校名替换）
├─ tools/  docs/
└─ TOOLS.md SKILL.md index.md
```

## 六、待完善（TODO）

- [ ] `frame_spec.py` 的 `diag / arch_bar / hex_bar / ring` 仍是占位（custGeom 未实现）
- [ ] 底卡分型 `corner / duo / topbar` 尚未在示范稿补页
- [x] 蒙版浓度经验值表见下
- [x] 数据块、时间轴、引文卡已封装进 `text-tools/tools/layouts.py`
- [x] 时间轴已封装（`layouts.timeline`）


---

## 七、蒙版浓度参考（2026-11-14）

| 背景明暗 | 蒙版色 | 不透明度 | 文字色 |
|---|---|---|---|
| 亮（浅色照片/白底） | 白 FFFFFF | 35% / 45%（默认）/ 55% | 深色 1F2E33 |
| 暗（深色照片/夜景） | 黑 12212E | 40% / 50%（默认）/ 60% | 白色 FFFFFF |
| 中等 | 白 | 50% + 文字加阴影 | 深色 |

规则：蒙版尺寸=卡片尺寸（略小 2–3px），只覆盖卡内，不外溢。
## 出画形状与质量门（2026-11-27 实测修正）

不少效果靠「形状整块停在画布外等命」（上滚、待命、入场）实现。质量门两种跑法规则不同：

| 跑法 | 出画文字 |
|---|---|
| `--roundtrip`（本库正常用法：参考件 → SVG → 改 → 回 PPTX） | **放行**，只查改字后的文字容量 |
| 默认 / 创作模式（从零画页面） | **ERROR**：版面契约禁止越出 viewBox |

所以按本库流程制作时，出画不需要任何标记。只有「从零创作」路径才可能要豁免：

```xml
<g id="stage-bottom" data-pptx-morph-staging="true" data-pptx-bounds="x y w h">
```

bounds 必须与画布完全不相交、且包住组内内容（容差 5%）；部分出画无法豁免。
工具：`python tools/stage_offscreen.py <svg目录|文件> [--check]`。
