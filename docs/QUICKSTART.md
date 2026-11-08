# 从 GitHub 到用起来：完整流程

> 适用于 Windows / macOS。全程约 5 分钟。
> 只想拿参考件改图、不装 DSH 的话，跳到 §5。

---

## 0. 前提检查

| 需要 | 怎么确认 |
|---|---|
| DeepSeek Harness | 能打开对话窗口即可 |
| Python 3.10+ | PowerShell / 终端里跑 `python --version`；Windows 若弹应用商店，用 `py -3 --version` 试 |
| 网络 | 能打开 github.com |

---

## 1. 下载（约 1 分钟）

1. 浏览器打开 **https://github.com/HY-hue113/ppt-effects**
2. 右侧点 **Releases**
3. 点最新版本 **v1.0.0**
4. 下载附件 **`ppt-effects-release.zip`**（27.4 MB）

> **不需要 GitHub 账号，不需要 git。** 这一步劝退的人最多，其实点两下就下完了。
> 下载慢的话，用手机热点，或换个时段重试；仓库页打不开就找作者要 zip。

**检查点**：下载下来是一个约 27 MB 的 zip，不是网页。

---

## 2. 安装到 DSH（约 2 分钟）

### 2.1 找到技能目录

按 `Win+R`，输入：

```
%USERPROFILE%\.dsh\skills
```

回车，会打开 `C:\Users\<你的用户名>\.dsh\skills`（`.dsh` 是隐藏文件夹，用这个方式最省事）。

macOS 是 `~/.dsh/skills`。

### 2.2 解压成 `ppt-effects` 文件夹

把 zip 里的内容解压到该目录，最终必须是这样：

```
skills\
└─ ppt-effects\
   ├─ SKILL.md        ← 必须在这一层
   ├─ index.md
   ├─ docs\
   ├─ effects\
   └─ tools\
```

**最常见的错**：解压出 `skills\ppt-effects\ppt-effects-release\SKILL.md`（多套了一层）。
处理：进到里层，把 `SKILL.md`、`docs`、`effects` 等**全部剪切到 `ppt-effects\` 下**。

### 2.3 装依赖

打开 PowerShell（或终端），执行：

```bash
pip install python-pptx Pillow lxml
```

装完不报错即可。若提示 `pip 不是命令`，换成 `py -3 -m pip install python-pptx Pillow lxml`。

### 2.4 重启 DSH

**技能目录只在 DSH 启动时扫描**，所以必须完全退出再打开（不是关窗口，是退出程序）。

**检查点**：重启后，`skills` 目录下有 `ppt-effects\SKILL.md`。

---

## 3. 验证装好了没（约 30 秒）

在 DSH 新会话里问一句：

> 效果6 是啥

**应该看到**：一句说明（单页、深色底、七段指针）+ 一张预览图 `docs/preview/效果6_加载指针.png` + 素材要求（1 张深色图）。

**如果它说"没听过这个效果"**：技能没扫到 → 回 2.2 查层级、2.4 查重启。

---

## 4. 正式使用

### 4.1 选效果

问它：

> 这些效果有什么区别？

它会给出 10 个效果的总览表（代号 / 一句话 / 需要几张图）。或者说"效果1 到 10 都是啥"。

### 4.2 带素材下单

> 用效果6 做一份，素材是这张图（附上图片），标题「国门」

它会：

1. 先跑 `check_materials.py` 校验素材（数量、画幅、明暗），
2. 再按 `effects/006_加载指针/recipe.md` 的管线出稿，
3. 渲染核对后给你 pptx。

**素材不合格时，它会先告诉你**（比如"这张图太亮，白字会看不清，建议换图或改白字"），而不是硬做。

### 4.3 想自己动手

打开 `effects\<效果>\recipe.md` 照做；`docs/USER-GUIDE.md` 里有完整说明。

---

## 5. 不装 DSH 也能用

1. 解压 zip；
2. 进 `effects\` 挑一个效果，打开 `reference\` 里的 pptx；
3. 右键图片 →「更改图片」换自己的图 → 按 F5 看效果；
4. 文字直接改（注意先确认字号，别把框撑破）。

要看更细的（素材规则、字体、常见问题），读 `docs/USER-GUIDE.md`。

---

## 6. 四个常见坑

| 现象 | 原因 / 处理 |
|---|---|
| 问"效果6 是啥"它说不知道 | 目录多套了一层 / 没重启 DSH |
| 报 `No module named pptx` | 依赖没装，回 2.3 |
| `pip` 不是命令 | 用 `py -3 -m pip ...` |
| 动画不动、只有淡入淡出 | PowerPoint 低于 2019（morph 需要 `p159`） |

---

## 7. 速查

```
仓库        https://github.com/HY-hue113/ppt-effects
下载        Releases → v1.0.0 → ppt-effects-release.zip
技能目录    %USERPROFILE%\.dsh\skills\ppt-effects\
依赖        pip install python-pptx Pillow lxml
验证话术    效果6 是啥
使用话术    用效果6 做一份，素材是这张图，标题「国门」
```
