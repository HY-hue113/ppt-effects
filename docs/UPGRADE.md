# 安装与升级

> Windows 为例；macOS 把 `%USERPROFILE%\.dsh` 换成 `~/.dsh`。
> 只想拿参考件改图、不装 DSH 的话，直接看 §5。

---

## 0. 先确认你属于哪种

| 情况 | 跳转 |
|---|---|
| 第一次装 | §1（zip）或 §2（git） |
| 装过旧版，想升级 | §3（zip 装的）或 §4（git 装的） |
| 只想要 pptx 改图 | §5 |

---

## 1. 新用户 · zip 安装

### 1.1 下载

1. 打开 `https://github.com/HY-hue113/ppt-effects`
2. 右侧点 **Releases**
3. 点最新版本（当前 `v1.1.0`）
4. **Assets** 里下载 `ppt-effects-v1.1.0.zip`（约 37 MB）

> **不需要 GitHub 账号，不需要 git。**
> 如果附件没传上去，就下载 `Source code (zip)`——内容一样，多了 `.github` 配置。

### 1.2 解压到技能目录

按 `Win + R`，输入：

```
%USERPROFILE%\.dsh\skills
```

回车，打开 `C:\Users\<你>\.dsh\skills`（`.dsh` 是隐藏目录，用这方式最省事）。

把 zip 解压到这里，文件夹命名为 **`ppt-effects`**。

**检查点**：

```
skills\ppt-effects\
    ├─ SKILL.md        ← 必须直接在这一层
    ├─ index.md
    ├─ README.md
    ├─ docs\
    ├─ effects\
    └─ tools\
```

**常见错误**：

| 错误形态 | 处理 |
|---|---|
| `skills\ppt-effects\ppt-effects-release\SKILL.md`（多套一层） | 进里层，把内容全部剪切到 `ppt-effects\` 下 |
| 只有 `docs/effects/tools`，没有 `SKILL.md` | 说明只解压了子目录，重新解压整包 |

### 1.3 装依赖

```powershell
pip install python-pptx Pillow lxml
```

- 报 `pip 不是命令` → `py -3 -m pip install python-pptx Pillow lxml`
- 弹微软应用商店 → `python` 是假入口，改用 `py -3`

### 1.4 重启 DSH

**完全退出**（托盘右键退出，不是关窗口）再打开。技能目录只在启动时扫描。

### 1.5 验证

新会话里问：

> 效果13 是啥

**成功**：给出说明（5 页、1–5 张图）+ 预览图 + 素材要求。
**失败**：“没听过这个效果” → 回 1.2 查层级、1.4 查重启。

---

## 2. 新用户 · git 安装

适合想“以后一键更新”的人。

```powershell
git --version                 # 没有就去 git-scm.com 装
cd $env:USERPROFILE\.dsh\skills
git clone https://github.com/HY-hue113/ppt-effects.git ppt-effects
pip install python-pptx Pillow lxml
```

然后重启 DSH。以后更新：

```powershell
cd $env:USERPROFILE\.dsh\skills\ppt-effects
git pull
```

---

## 3. 老用户 · 从旧版升级（zip 装的）

### 3.1 备份自己改过的东西

没改过就跳过。改过的通常是 `effects\<效果>\reference\*.pptx`。

```powershell
Copy-Item "$env:USERPROFILE\.dsh\skills\ppt-effects" `
          "$env:USERPROFILE\Desktop\ppt-effects-备份" -Recurse
```

### 3.2 退出 DSH

文件被占用时覆盖会失败。

### 3.3 整体覆盖

1. 删除或改名旧的 `skills\ppt-effects`
2. 新包解压成 `skills\ppt-effects`
3. 把 3.1 备份里改过的文件拷回去
4. 重启 DSH

**必须整体覆盖**，只换一部分会坏：

| 本次改动 | 只覆盖一部分的后果 |
|---|---|
| 预览图 `png` → `jpg` | 只换 `docs/preview/` 不换 `spec.json`，预览图路径全失效 |
| 新增 `effects/011–015` | 新效果没进来，还是 10 个 |
| `check_materials.py` 加 `--brief` | 工具没换，传参报错 |

### 3.4 验证

```powershell
python tools\check_materials.py --list
```

应列出 **15 个效果**；只有 10 个 = 没覆盖成功。

---

## 4. 老用户 · 升级（git 装的）

```powershell
cd $env:USERPROFILE\.dsh\skills\ppt-effects
git pull
```

看到 `Updating xxx..yyy` 即成功，然后重启 DSH。

**报冲突**（`Your local changes would be overwritten`）说明你改过库内文件：

```powershell
git stash
git pull
git stash pop     # 有冲突就手动合并
```

---

## 5. 不装 DSH 也能用

1. 解压 zip
2. 进 `effects\` 挑效果，打开 `reference\` 里的 pptx
3. 右键图片 →「更改图片」换图，按 **F5** 看动效
4. 文字直接改

更细的规则见 `docs/USER-GUIDE.md`。

---

## 6. 升级后自查

| 检查 | 动作 | 期望 |
|---|---|---|
| 效果数量 | `python tools\check_materials.py --list` | 15 行 |
| 技能识别 | 问「效果13 是啥」 | 答出“1–5 张” |
| 参考件能开 | 打开 `effects\011_中式滑转\reference\效果11-多图版.pptx` | 10 页正常 |
| 预览图 | 看 `docs\preview\` | 15 个 `.jpg` |

---

## 7. 报错对照

| 报错 | 原因 | 处理 |
|---|---|---|
| 问效果说“没听过” | 层级错 / 没重启 | 检查 `SKILL.md` 位置 |
| `No module named pptx` | 依赖没装 | `pip install python-pptx Pillow lxml` |
| `pip 不是命令` | PATH 无 pip | `py -3 -m pip install ...` |
| 覆盖时“文件正被使用” | DSH 还开着 | 完全退出 DSH |
| 动画只有淡入淡出 | PowerPoint < 2019 | 升级 Office（“平滑”切换需要 2019+） |
| 字体变了 | 参考件不嵌字体 | 装霞鹜文楷 / 思源黑体 |

---

## 8. 速查

```
仓库        https://github.com/HY-hue113/ppt-effects
下载        Releases → 最新版 → ppt-effects-vX.Y.Z.zip
技能目录    %USERPROFILE%\.dsh\skills\ppt-effects\
依赖        pip install python-pptx Pillow lxml
验证        python tools\check_materials.py --list
```
