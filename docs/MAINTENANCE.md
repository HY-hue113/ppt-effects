# 修改与更新操作说明

> 面向维护者：新增效果、改已有效果、改脚本、发布前自查。
> 使用者请读 `USER-GUIDE.md`。

---

## 1. 目录与命名约定

```
effects/001_湖景斜切/         三位序号_短名，只增不改名
    recipe.md                人读；结构表 / 参数 / 坑 / 复用管线
    spec.json                机器读；唯一事实来源
    reference/*.pptx         参考件（可多个：短版/长版/变体）
    svg/                     可选：直接可渲染的模板
    animations.json          可选：对象动画配置
tools/*.py                   一次性脚本也放这里，头部注明用途
index.md                     索引，按创建顺序追加
```

**硬规则**：`index.md` 的代号、目录名、`spec.id` 三者必须一致；序号不复用、不插空。

---

## 2. 新增一个效果

1. **先有可直接打开的参考件**（手工在 PowerPoint 里调好的那版），改名放进 `effects/NNN_xxx/reference/`。
2. 写 `spec.json`：几何用 EMU、字体、切换、以及 `materials` 约束（见 §4）。
3. 写 `recipe.md`：结构表 + 关键参数表 + 制作管线 + **踩过的坑**（每条都要能指向具体现象）。
4. 若可直接用 SVG 生成，放 `svg/` 模板；否则在 recipe 里写明“模板移植”。
5. `index.md` 追加一行：代号 / 名称 / 素材要求 / 动画要点 / 切换 / 日期 / 目录。
6. 用新素材实跑一遍，跑通再入库。

---

## 3. 修改已有效果

顺序固定，不要跳：

1. **先改参考件**（PowerPoint 里手工调好）；
2. **再把几何/参数抄进 `spec.json`**（EMU 原值，别换算成 cm 再换算回去）；
3. **同步 `recipe.md`** 的表和坑；
4. **同步 `index.md`** 该行的“素材要求 / 动画要点”；
5. 用 `tools/check_materials.py <效果> <图>` 回归一次；
6. 用参考件重跑一遍管线，确认渲染无差。

> 只改 doc 不改 spec、或只改 spec 不改参考件，都会让下一个使用者踩坑。

---

## 4. `spec.json` 字段速查

| 字段 | 作用 |
|---|---|
| `canvas` | 画布（px / cm / EMU），必须与参考件一致 |
| `layers[]` / `frames` / `photo_slots` | 各效果的几何，按效果类型不同 |
| `transition` | 切换类型、时长、触发方式；**morph 必须是 `byObject`** |
| `materials` | 素材约束，见下 |
| `fonts` | `used_in_reference` / `recommended_free` |

### materials 子字段

| 键 | 含义 |
|---|---|
| `min` / `max` | 必需张数 |
| `recommended` | 推荐张数（不足只警告） |
| `aspects` / `aspect_tolerance` | 槽位画幅与容差 |
| `orientation` | landscape / portrait / mixed / any |
| `deficit` | 缺图策略：`none` 阻断、`repeat` 重复用图、`fewer-pages` 自动少页、`keep-old` 沿用旧图（慎用） |
| `crop_allowed` | 允许居中裁切（不判为拉伸） |
| `checks.max_mean_luma` | 上限亮度（深色底效果） |
| `checks.min_white_border` | 白底阈值；配 `white_bg_blocking` 决定是警告还是阻断 |
| `checks.accept_alpha` | 允许“已抠好的透明 PNG”替代白底 |
| `checks.text_contrast` | 文字对比：`title_color` + `min_mean_luma`，黑字配亮图 |

---

## 5. 参考件制作规范（对外发布用）

参考件是最大的风险载体，按这套做：

1. **素材自有**：照片用自己拍的；他人的模板/插画/配乐一律替换或删除。
2. **清文档属性**：`docProps/core.xml` 的 `dc:creator`、`cp:lastModifiedBy` 改成库名，清掉邮箱。
3. **剥嵌入字体**：删 `ppt/fonts/`、`presentation.xml` 的 `embeddedFontLst` 与对应 rels、Content_Types。
4. **删音视频**：删媒体部件前，先删 `<p:audio>`、`a:audioFile`、`p14:media` 与 rels，否则残留引用。
5. **清孤儿媒体**：只保留被 rels 引用的；顺手删 `notesSlides` 里指向已删幻灯片的悬空页。
6. **压图**：长边 ≤1280、JPEG q85；照片别存 PNG（一张能到 2.5 MB）。
7. **保留编号**：清文字做纯版式时，保留 `01–04`、`55`、`PART01–05`、`·`、`◆` 这类结构字符，否则版式塌掉。

---

## 6. 脚本改动约定

- **判定阈值进 spec，不进代码**：`check_materials.py` 只读 `materials`，不要在脚本里写死数字。
- 新增检查项时，同步更新 `README`/本文档的字段表。
- 脚本里不写本机绝对路径；用参数或相对路径。
- 改完跑 `python -m py_compile`，再拿两个真实案例回归（一 PASS 一 BLOCK）。

---

## 7. 发布前自查清单

```bash
# 1) 文本泄漏：绝对路径 / 邮箱 / 昵称 / 私有配置
# 2) 参考件：作者字段、嵌入字体数、rels 断链数
# 3) 音视频残留
# 4) 脚本可编译
# 5) 配方引用的 tools/svg/reference 是否都在
```

逐项要求：

- [ ] 所有 `.md/.py/.json` 无本机路径、无邮箱、无他人昵称
- [ ] 每个参考件：`dc:creator` 统一、嵌入字体 0、断链 0
- [ ] 无 mp3/wav/mp4 残留
- [ ] `index.md` 与目录一一对应，`spec.json` 全部可解析
- [ ] 每个效果都能用 PowerPoint 打开且页数正确
- [ ] `LICENSE` / `README` / `USER-GUIDE` / 本文件 齐全

---

## 8. 已知坑（都在具体 recipe 里有展开）

| 现象 | 根因 |
|---|---|
| 打开报“文件已损坏” | 删了形状但 `<p:timing>` 仍引用 spid |
| 同样报损坏 | `notesSlides` 指向已删 slide |
| morph 失效 | `option="object"` 不是合法枚举，必须 `byObject` |
| 亮窗消失/多出一个亮圆 | 转盘是 `useBgFill=1`（幻灯片背景填充），不能改 `noFill` |
| 文字叠在标题上 | 滚屏类效果没有把上一段整体移出画布 |
| 玻璃变彩色菱形 | 图文框必须 `useBgFill` + 幻灯片背景，不能填图片 |

---

## 9. 版本记录格式

`RELEASE-NOTES.md` 每次发布追加一节：

```markdown
## vX.Y · YYYY-MM-DD
- 新增：效果N（一句话）
- 修改：效果N 的 X（改了什么、为什么）
- 修复：具体 bug
- 兼容性：是否需要重跑已有项目
```
