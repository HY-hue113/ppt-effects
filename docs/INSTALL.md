# 安装到另一台电脑

> 目标：在另一台机器的 DeepSeek Harness 上把它当 Skill 用起来。

## 0. 先看要拷什么

两种都行：

| 方式 | 拷什么 | 适用 |
|---|---|---|
| **A. 直接搬库**（推荐） | 整个 `ppt-effects` 目录（含 `docs/`、`effects/`、`tools/`） | 想连工作库一起带走 |
| **B. 用发布包** | `ppt-effects-release` 目录或它的 zip | 只要可公开的那一版 |

两者结构相同，`skills/ppt-effects/SKILL.md` 是入口。

## 1. 放到 Skill 目录

DeepSeek Harness 从用户目录下的 `skills/` 发现技能：

```
C:\Users\<你的用户名>\.dsh\skills\ppt-effects\
    SKILL.md        ← 必须有，frontmatter 里的 name 就是调用名
    index.md
    docs/
    effects/
    tools/
```

把整个目录拷到 `\.dsh\skills\ppt-effects\` 即可（目录名建议就叫 `ppt-effects`，与 SKILL.md 的 `name` 一致）。

macOS / Linux 对应 `~/.dsh/skills/ppt-effects/`。

## 2. 装依赖

```bash
pip install python-pptx Pillow lxml
```

验证：

```bash
python tools/check_materials.py --list
```

能打印出 10 个效果的素材要求，说明依赖和文件结构都对了。

## 3. 字体（重要）

参考件**不含嵌入字体**，字形取决于本机已装字体。建议先装：

- 霞鹜文楷 LXGW WenKai（替代华文行楷/楷体）
- 思源黑体 Source Han Sans、思源宋体 Source Han Serif
- HarmonyOS Sans SC 或 Noto Sans SC

不装也能跑，只是回退到系统默认字，观感会变。

## 4. 可选：ppt-master

只有**效果1、效果2 短版**的 SVG → PPTX 管线需要它（脚本 `project_manager.py`、`svg_to_pptx.py`、`svg_quality_checker.py`）。
其余 8 个效果是“参考件 + 补丁”路线，不需要。

## 5. 首次自检

```bash
# 1) 素材检查器能不能跑
python tools/check_materials.py --list

# 2) 拿一张图试判定
python tools/check_materials.py 效果1 你的图.jpg

# 3) spec 能否解析
python -c "import json,glob;[json.load(open(p,encoding='utf-8')) for p in glob.glob('effects/*/spec.json')];print('ok')"
```

## 6. 在会话里调用

在 DSH 里直接说：

> 用效果1 做一份，素材是这张图，标题「雪山」

Agent 会读 `SKILL.md` → `index.md` → 该效果的 `recipe.md`/`spec.json`，然后跑 `check_materials.py` 校验素材。

## 7. 已知迁移问题

| 问题 | 原因 / 处理 |
|---|---|
| 脚本报路径不存在 | 发布包里的路径已打码成 `<ABS_PATH>` / `<LIB_ROOT>`；`check_materials.py`、`patch_effect.py`、`patch_effect2.py` 是参数化的，可直接用；`build_effect*.py` 是一次性构建脚本，需按本机路径改 |
| 提示缺 python-pptx | 装依赖；注意有的机器 `python` 指向 Windows 商店的假入口，用 `py -3` 或真实解释器路径 |
| 参考件打不开 | 需要 PowerPoint 2019+（morph 切换 `p159`）或 LibreOffice；旧版 PowerPoint 会降级成淡入淡出 |
| 中文字体发虚/回退 | 见 §3 |
| 找不到 skill | 确认目录是 `\.dsh\skills\ppt-effects\SKILL.md`，且 frontmatter 有 `name: ppt-effects` |

## 8. 建议的第一次实验

1. `check_materials.py --list` 跑通；
2. 挑**效果6**（单页、只要一张深色图）做最小验证：拷参考件 → 换 `ppt/media/image1.jpeg` → 打开；
3. 再试**效果1**（走 SVG 管线，依赖最多，能跑通说明环境齐了）。
