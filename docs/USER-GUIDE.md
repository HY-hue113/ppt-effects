# 用户使用说明

> ppt-effects 是一套「PPT 动效配方库」：每个效果是一份 recipe（人读）+ 一份 spec.json（机器读）+ 一份参考件（可直接打开）。
> 本说明面向使用者；修改库本身请看 `MAINTENANCE.md`。

---

## 1. 它能做什么

- 你有一张/几张照片和一段文案，想要某种已经调好的动画效果（斜切滑入、玻璃散开、环形开场……）
- 不用从零调动画：按 recipe 执行，几何、时序、切换参数都写在 spec 里
- 10 个效果的产物都是可编辑的原生 PPTX，不是视频

## 2. 环境准备

| 需要 | 说明 |
|---|---|
| Python 3.10+ | 依赖 `python-pptx`、`Pillow`、`lxml` |
| PowerPoint | 用于预览、放映、最终校对（Linux 可用 LibreOffice 近似） |
| ppt-master（可选） | 效果1/2短 的 SVG→PPTX 管线依赖它；其余效果是「参考件 + 补丁」路线，不需要 |
| 字体 | 参考件**不含嵌入字体**，字形取决于你机器上装的字体。优先用开源字体（见 §6） |

```bash
pip install python-pptx Pillow lxml
```

## 3. 目录结构

```
index.md                效果目录：代号 → 目录 → 素材要求 → 动画要点
SKILL.md                作为 Agent Skill 使用时的入口（可让 AI 直接读）
effects/NNN_xxx/
    recipe.md           人读：结构、参数、坑、复用方法
    spec.json           机器读：精确几何/字体/切换/material 约束
    reference/*.pptx    参考件：可打开、可当模板套
    svg/                仅 001、002 有：可直接走 SVG 管线
tools/
    check_materials.py  素材校验（数量/画幅/明暗/白底）
    patch_effect*.py    导出后精确校正
```

## 4. 三步走

### 第 1 步：选效果

打开 `index.md`，按「素材要求」筛：

| 代号 | 名称 | 需要几张图 |
|---|---|---|
| 效果1 | 湖景斜切滑入 | 1 张 4:3，不能太暗 |
| 效果2 | 玻璃碎片散开 | 短版 1 张 / 长版 1–5 张 |
| 效果3 | 哲学字动 | 1–5 张（推荐 4–5） |
| 效果4 | 扇形展开 | 1 张 ≈1.77 |
| 效果5 | 三段上滚 | 1 张 |
| 效果6 | 加载指针 | 1 张深色图 |
| 效果7 | 环形开场 | 1 张（整套另有 13 处媒体槽） |
| 效果8 | 纯矢量动画 | 不用图，只换文字 |
| 效果9 | 金边弧线 | 16 个槽（可重复用图） |
| 效果10 | 暗红塔影 | 白底主体图**或**已抠好的透明 PNG |

### 第 2 步：校验素材（必做）

```bash
python tools/check_materials.py 效果1 照片.jpg
```

输出三选一：

- `PASS`（退出码 0）：可以开工
- `WARN`（退出码 0）：能开工，但**必须先跟需求方确认**列出来的问题
- `BLOCK`（退出码 1）：不要硬做，先补素材

它检查五件事：数量、画幅（能否裁切/会不会拉伸）、方向、明暗（文字对比）、以及特定前提（白底/透明）。

### 第 3 步：按 recipe 出稿

打开 `effects/<效果>/recipe.md`，照「制作管线」一节执行。以效果1 为例：

```bash
# 1) 建项目（quick 模式）
python scripts/project_manager.py init 我的封面 --format ppt169 --quick-generate

# 2) 按 recipe 处理素材放进 images/，两页 SVG 写终态

# 3) 质检 + 导出
python scripts/svg_quality_checker.py <项目> --quick-generate --canonical-authoring --stage final --json
python scripts/svg_to_pptx.py <项目> -f ppt169 --quick-generate -t none -o 导出.pptx

# 4) 精确校正（把对象移到 spec 指定的起止位置、补阴影、写切换）
python tools/patch_effect.py --in 导出.pptx --spec effects/001_湖景斜切/spec.json --out 定稿.pptx --text 标题 --variant short
```

其余效果的管线见各自 recipe：多为「复制参考件 → 换 media → 改文字」，几何一律不动。

## 5. 素材规则（总结）

1. **同一个效果别凭印象套**：先跑 `check_materials.py`，判定阈值写在 spec 的 `materials.checks` 里。
2. **画幅优先“裁”不“拉”**：spec 的 `crop_allowed: true` 表示允许居中裁切；没有这个字段就按原比例给图。
3. **深色效果（效果6）要暗图**；**黑字效果（效果1）要亮图**，亮度 <140 会触发警告。
4. **白底效果（效果10）**：给白底原图，或直接给抠好的透明 PNG —— 后者跳过反推，最快。
5. **可以重复用图**：效果9 只有 4 张也能铺 16 槽（`deficit: repeat`），不会混入参考件的旧图。

## 6. 文案与字体

- **文本框四问**：数量、每段内容、字体、字号——不要用库里默认值直接代替确认。
- **推荐开源字体**（可商用，无需授权）：

| 想要的感觉 | 用 |
|---|---|
| 手写/楷体 | 霞鹜文楷 LXGW WenKai |
| 黑体标题 | 思源黑体 Source Han Sans / MiSans |
| 宋体正文 | 思源宋体 Source Han Serif / Noto Serif SC |
| 系统感 | HarmonyOS Sans SC / Noto Sans SC |
| 西文衬线 | EB Garamond / Liberation Serif |

商业字库（华文、方正、汉仪等）个人非商用可用，**对外发布或商用须另行取得授权**。

## 7. 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| PowerPoint 提示“文件已损坏” | 多半是删了形状但 `<p:timing>` 还引用它的 spid；删形状时要同步清 timing |
| 平滑切换变成淡入淡出 | 两页对应对象**名字不一致**；morph 靠名字配对 |
| 字体在别人机器上变了 | 参考件不嵌字体；要么让对方装字体，要么换开源字体 |
| 效果6 的第 1 页全黑 | 正常，刀片/指针都是“按 F5 才动”的首帧状态 |
| 效果8 第 33 页字叠在一起 | 同上，静态帧是动画中间态 |
| 效果9 少给图会不会混入旧图 | 不会；`deficit: repeat` 是重复用你给的图 |

## 8. 发布前必须自查

如果你要把产出发给别人或商用：

- 照片里的人像（肖像权）是否可辨认、是否获授权；
- 字体是否已换成开源字体；
- 引用他人模板/插画/音乐的部分是否已替换。

`check_materials.py` 只管素材规格，不管这些。
