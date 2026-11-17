# 图片处理区 · Image Tools

> 跨效果复用的**图片处理工具集**。调用任意效果时，如果要用图，可以直接调用本区的脚本与规范，不必每次重写。

## 一、这里有什么

| 文件 | 作用 |
|---|---|
| `tools/build_demo10.py` | 示范稿唯一构建脚本（旧页拷入 + 新页生成 + 一次性导出） |
| `tools/mask_lib.py` | 生成"形状遮罩图"（透明 PNG）：切角 / 圆形 / 六边 / 拱形 / 斜切 |
| `demo/reference/图片取景形状_示范.pptx` | **唯一示范稿，10 页**（5 旧形状 + 5 新形状 + 7 组合） |
| `demo/preview/p01–p10.png` | 10 页示范截图 |
| `demo/技法说明.md` | 参数、纪律、扩展方式（原效果18 的配方） |
| 本文件 | 总览、调用方式、与其他效果的关系 |

## 二、五种基本形状

| 形状 | 轮廓 | 适用 |
|---|---|---|
| `cut` **切角** | 左上 18% 切角 | 稳、正式，方案/论文 |
| `circle` **圆形** | 正圆 | 人像、特写、柔和 |
| `hex` **六边** | 六边形 | 现代、有节奏 |
| `arch` **拱形** | 上圆下方（门洞） | 建筑、门、"入口"意象 |
| `diag` **斜切** | 平行四边形 | 动势、前后对照 |

做法：先生成与形状**同轮廓的透明 PNG**（PIL 遮罩），再把**同形状描边**叠在图上。
这样在 PPT 里"图是图片、框是形状"，两者都能单独换。

### 五种新增形状（v2）

| 形状 | 轮廓 | 适用 |
|---|---|---|
| `rhombus` **菱形** | 四角顶点 | 强调、指示、单点聚焦 |
| `capsule` **胶囊** | 两端全圆 | 标签、按钮感、人物剪影 |
| `wave` **波浪边** | 下边缘正弦波 | 水、风、柔性过渡 |
| `ticket` **票券齿孔** | 左右各一个半圆缺口 | 入场券、门票、纪念卡 |
| `broken` **破格矩形** | 右下角缺一块 | 打破网格、留白、透气 |

## 三、参数（已调好，直接用）

| 项 | 值 |
|---|---|
| 切角量 | 短边 × 18% |
| 六边斜边 | 短边 × 22% |
| 拱形圆角 | 短边 × 45%（只圆上两角） |
| 斜切偏移 | 宽 × 20% |
| 描边 | 2–3px 强调色；叠压时后图用白描边 6px 分层 |
| 图间距 | ≥ 宽 × 12%（小于会挤） |

## 四、组合范式

| 组合 | 规则 |
|---|---|
| 同形同尺寸并列 | 5 张 200×200，间距 = 宽 16% |
| 同形不同尺寸 | 主图 380 + 辅图 180×2，靠尺寸分主次 |
| 异形叠压 | 六边 420 + 圆 240 压角，后图加 6px 白描边 |
| 多图阵列 | 1 大 420×300 + 4 小 190×140，同形同距 |
| 异形对置 | 斜切 460 + 拱形 380，中间留 150px 竖缝 |
| 同形阶梯 | 切角 380→300→220 沿对角线递减 |
| 三形咬合 | 圆 380 + 六边 340 + 切角 300，相邻压边 40px，后叠加 8px 白描边 |
| 网格 + 异形破格 | 3×3 切角网格（160/间距 24），中心换菱形 240 放大 1.5 倍溢出格线 |
| 同图多形状并置 | 同一张图裁 4 种形状并排（220px），选型时最直观 |

## 五、纪律

1. **一页最多两种形状**，且一主一辅
2. 同页多图必须**同形同尺寸**（主图除外）
3. 叠压时后图必须加**白描边**，否则两张图边界糊在一起
4. 拱形只用于竖向构图，横图塞进去会裁掉太多

## 六、怎么调用

```python
import sys
sys.path.insert(0, r'<LIB_ROOT>\image-tools\tools')
from mask_lib import shaped_image, outline_shape, paste_framed

# 1) 生成遮罩图（透明 PNG）
shaped_image('原图.jpg', '输出.png', w=420, h=300, shape='hex', fit='cover')

# 2) 生成同轮廓描边（几何同源，不会错位）
outline_shape('描边.png', w=420, h=300, shape='hex', width=3, color='#4472C4')

# 3) 在效果稿里：图在前、描边在后
```

具体函数签名见 `tools/mask_lib.py`。

## 七、可扩展

新增形状只需在 `shaped_image()` 里加一个分支：菱形、胶囊、波浪边、票券齿孔……

## 八、与效果库的关系

- 本区**不绑定任何效果**，是公共设施
- 效果18（图片取景形状）是本区的"示范应用"
- 调用效果 1–17 时，只要涉及图片，都可以用本区处理图，再放进效果版式

## 九、维护纪律（重要）

**示范稿只有一份**（`demo/reference/图片取景形状_示范.pptx`），当前 10 页。新增形状或组合时：

1. 在 `tools/build_demo10.py` 里加页（旧页自动从源目录拷入）
2. 重跑脚本 → 覆盖这一份 pptx
3. 重出 `demo/preview/pNN.png`
4. README / 技法说明补参数

**不要再产出第二份示范 pptx。**

**只要新增了图片处理方式（新形状、新组合、新参数），必须同步更新示范 PPT。**

要更新的东西：

1. `tools/mask_lib.py` — 加形状分支（`mask()` / `outline_svg()`）
2. `tools/build_extra_demo.py` 或新建构建脚本 — 加对应演示页
3. `demo/reference/` — 重新导出示范 pptx
4. `demo/preview/` — 重新渲染预览图
5. 本 README 与 `demo/技法说明.md` — 补参数与用法

原因：示范稿是这个工具区唯一的"可视化说明书"，不同步就等于工具没文档。

## 七、边框 = 形状 + 图片填充（2026-11-14 定稿）

**禁止再把"裁好形状的图"贴到页面上。** 图片必须作为**形状的填充**存在：

```xml
<p:sp>
  <p:spPr>
    <a:prstGeom prst="snip2SameRect">…</a:prstGeom>   <!-- 或 a:custGeom -->
    <a:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
  </p:spPr>
</p:sp>
```

好处：圆角可拖、图可换、描边可改、顶点可编辑——不再需要重跑脚本。

### 边框 → 原生形状映射表（`tools/frame_spec.py`）

| 边框 | 实现 | 说明 |
|---|---|---|
| `cut` 切角 | `prstGeom snip2SameRect` adj1=12000 | 预设形状，角度可调 |
| `arch` 拱形 | `prstGeom round2SameRect` adj1=40000 | 圆角上、直角下 |
| `hex` 六边 | `prstGeom hexagon` | |
| `circle` 圆 | `prstGeom ellipse` | |
| `rhombus` 菱形 | `prstGeom diamond` | |
| `capsule` 胶囊 | `prstGeom roundRect` adj=50000 | 全圆角 |
| `diag` 斜切 | `custGeom` | 矩形去两角 |
| `duo_circle` 双圆 | `custGeom` **并集路径** | 交点断开、只走外侧弧（已实测） |
| `arch_bar` / `hex_bar` | `custGeom` 并集 | 组合边框 |
| `ring` 镂空 | `custGeom` 外矩形−内矩形 | 导航镂空同此思路 |

### 合并形状算子（对应 PowerPoint「合并形状」）

| 算子 | 用途 |
|---|---|
| 结合 Union | 双圆取景、L 形卡 |
| 相交 Intersect | 图章式取景 |
| 剪除 Subtract | 导航镂空、卡片挖角 |
| 拆分 Fragment | 一块形状分区上色 |
| 组合 Group | 多形状整体移动 |
| 编辑顶点 | 自由造型（custGeom 可拖点） |

### 调用

```python
import sys; sys.path.insert(0, r'<LIB_ROOT>\image-tools\tools')
from frame_spec import sp_xml
xml = sp_xml(12, "PicFrame", x=920, y=240, w=640, h=420, frame="cut", rid="rId7")
```

## 八、custGeom 三条硬要求（2026-11-14）

custGeom **可以用**（PowerPoint 的「合并形状」「编辑顶点」产物就是它），
但手写时必须满足三条，否则 PowerPoint 会报需要修复或填充失效：

| # | 要求 | 症状 |
|---|---|---|
| 1 | <a:path ... fill=norm> | 缺了 → 形状有轮廓但填充为空 |
| 2 | spPr 顺序 xfrm → geom → 填充 → ln | 错序 → **提示需要修复** |
| 3 | 路径首尾闭合、坐标 ∈ [0,100000] | 越界/未闭合 → 修复或渲染错乱 |

校验脚本：	ools/custgeom_check.py

`python
from custgeom_check import check_custgeom, check_sppr_order
check_custgeom(geom_xml)     # fill / 坐标 / 闭合
check_sppr_order(sp_xml)     # 元素顺序
`

实测结论：
- P12 双圆第一次手写漏了 ill=norm → 圆里有轮廓无照片
- 第二次把 lipFill 追加在 ln 之后 → 提示需要修复
- 补齐两条后合并成功，一条金边、图片填充连续
