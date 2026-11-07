# 建仓库时要填的字段（可直接复制）

> 这份不进 README，纯粹给"新建仓库"那一步用。

## Description（仓库描述）

推荐（中文）：

```
把调好的 PPT 动效存成可复用配方：10 个效果，每个含 spec 与参考件，给图给字即可出同款动画
```

英文版（可放在中文后，或单独用于英文场景）：

```
Reusable recipes for PowerPoint motion effects — 10 effects with spec + reference deck
```

备选：

| 风格 | 文案 |
|---|---|
| 技术向 | PPT 动效配方库：recipe + spec.json + reference.pptx，附素材校验器与补丁脚本 |
| 极简 | 10 个可直接套用的 PPT 动效配方（含参考件与素材校验） |

**别写**："AI 一键生成"（没做这件事）、"免 PowerPoint"（放动画需要 PowerPoint 2019+ 或 LibreOffice）。

## Topics（最多 20 个，建议 8 个）

```
powerpoint  pptx  animation  morph  ppt-template  python-pptx  ai-skill  deepseek-harness
```

## Website

留空。以后有介绍文章再填。

## Release（v1.0.0）

- Tag：`v1.0.0`
- 标题：`v1.0.0 首个版本`
- 附件：`ppt-effects-release.zip`（27.4 MB）
- 说明：可直接用 `RELEASE-NOTES.md` 的内容

## 一句话转发文案

> 把 10 个调好的 PPT 动效做成了可复用配方，附参考件和素材校验器，给张图就能出同款动画。

## 推送命令（本地仓库已就绪）

```bash
cd D:\skilltest\release\repo
git remote add origin https://github.com/<用户名>/ppt-effects.git
git push -u origin main --tags
```

凭据弹窗由本机 Git Credential Manager 处理，不要用账号密码（GitHub 已停用）。
