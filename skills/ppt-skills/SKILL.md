---
name: ppt-skills
description: >
  Use when the user asks to 做PPT / 生成演示文稿 / 用我的风格做PPT / 把文档做成PPT /
  make a presentation, or mentions "ppt-skills" or the business-blue 风格. Personal-style,
  template-floored, editable-PPTX generation.
---

# PPT Skills

> 在 ppt-master 引擎之上，以**我人工打磨的模板为"质量地板"**，把**真实内容**适配进去，产出更接近成品的**可编辑** PPTX。
> **主线 = 原生 DrawingML 蒸馏**：模板里的母题 / 风格件已逐页蒸进 `assets/business-blue/` 素材库；生成 = **builder 调库素材组装 + 真实内容回流**，再 PowerPoint COM 渲染核对。

**核心流水线**：`源文档/主题 → 选风格(business-blue floor) → 内容拆页型 → 逐页:取页型配方 + 调 catalog 素材 → builder place() 组装 + 内容回流 → §5 自检 → COM 渲染核对 →（可选）入合集册`

---

> [!CAUTION]
> ## 🚨 全局执行纪律（最高优先级，违反即失败）
>
> 1. **地板优先、不自由设计；但配方是参考、非照抄** —— 默认用 business-blue 素材库当质量地板，锁画布 / 配色 / 字体 / 栅格；无合适素材才回退自由设计（并按 Step 7 蒸新件入库）。**`页型配方.md` 给的是母题 / 骨架 / 用哪件素材的"蓝图"，不是逐字答案**：照骨架搭，但**内容来自用户、版式按页 adapt**，**别让所有页 / 所有稿子长成一模一样**（同一 deck 里过渡页等可有意复用版式，其余别千篇一律）。
> 2. **🔴 只调库、禁内联** —— builder **只能 `place()` / `_load()` 调 `assets/business-blue/` 的素材**来组装；**严禁在 builder 里内联写死 substantial 图形 XML**。判据：图形若「换页还想用」或「用户会想换 / 填」= 资产，必须先蒸入库（`蒸馏流程.md`）；只有纯版式（textbox / 对齐 / 文字）才许现写。**别凭记忆近似画**模板里现成的好素材。
> 3. **🔴 字号刻度统一** —— 同角色跨页同字体（全库微软雅黑）同字号：内容页标题 / 引导 28 粗、正文 ≈14、签名页主标按模板实测放大、硬下限 14。**先量模板真值**（`sz`÷100），别靠手感（`design-language.md` §1.3）。
> 4. **🔴 文本别硬断词** —— 正文按模板真实自然段、让其**自动换行**，严禁手塞 `\n` 把词断开。
> 5. **内容契约回流** —— 真实内容长度 ≠ 模板：`折行 → 缩字(到下限) → 扩块 / 挪邻居 → 换更密页型`；文字不得越容器隐形边界、不得溢出 / 截断。
> 6. **占版铺满、严禁胆怯小字** —— 主视觉吃满安全区 ~50%，别缩中央留大白；字号不低于模板（记忆 `ppt-readable-scale`）。
> 7. **进料即净化** —— 导入的 pptx / 蒸出的素材先剔除公司 DLP 图章 + 水印（©方能演亦 / EagleCloudWatermark），绝不写入产物。
> 8. **增量不重建** —— 素材 / palette / preview / gallery **只动当前学 / 改的页**，绝不全量重渲（毁手工深底、费 token）；`render_truth.ps1` 已设护栏。
> 9. **设计感是硬指标** —— 生成**前**读 `design-language.md`（§1 原则 / §1.3 字号 / §四 出件）+ 该页型在 `页型配方.md` 的一节；生成**后**过 §5 自检逐条、任一 NO 回修；**严禁退化成纯色卡片网格**。
> 10. **渲染核对 = 真图** —— `render_slides.ps1`(COM) 渲当前页核对（所见即所得，**全反斜杠路径**）；E_FAIL 多是命名空间 / 主题坑（`design-language.md` §6.18/19/22/23/25），二分定位。

> [!IMPORTANT]
> ## 🔌 裁剪范围
> 本 skill **不含**：AI 生图、配音 / 旁白。聚焦「文档 / 主题 → 个人风格可编辑 pptx」。用户自带图片照常处理。

---

## 知识与规格（进入相应阶段前必读）

| 资源 | 用途 |
|---|---|
| **`references/design-language.md`** ⭐ | **生成必读·设计主脑**：景深 / 光影 / 占版 / **字号刻度** / 凸出 / 非扁平 原则 + §四 出件四选一 + **§五 自检清单** + **§六 避坑**。每页生成前读、生成后自检 |
| **`references/页型配方.md`** | **页型 cookbook**：15 种 content_type 的层叠配方（母题 / 骨架 / 用哪件素材）。**做某页型只查那一节**，不必通读 |
| **`references/资产地图.md`** | **素材在哪 + 怎么调**：`catalog.json` / `vectors`(注入) / `rasters`(图片) / `builders`(参数化) / `_palette`(看样) / `templates/icons`(≈7981 图标) |
| **`references/蒸馏流程.md`** | **学一页 / 扩库流程**（10 步蒸馏标准）。**仅当缺页型素材、要从模板学新页扩库时读**；日常生成不必读 |
| `assets/business-blue/` | **原生素材库**：`catalog.json`(权威清单) / vectors / rasters / builders / _palette / _preview |
| `_gallery/`（见其 README） | **成果**：合集册 `business-blue-master.pptx` + 每页 png |
| `../../docs/` | 方法论 / 技法库 / 美学门规格（立项、深究时读） |

---

## Workflow（原生蒸馏主线）

### Step 0 · 前置设计门（硬门，不可跳）
开工前先确认有**用户已批准的《逐页设计稿》**（由 `ppt-design` skill 产出）：
- **无** → 不进入制作，引导用户先走 `ppt-design`（设计前脑：逐页四问 → 出设计稿 → 审批）。
- **有** → 严格按设计稿逐页制作（设计稿的"表达式 / 真实素材"即本次制作依据，**不在制作阶段另起设计**）。

### Step 1 · 源内容处理
非 Markdown 先转：`python3 ${SKILL_DIR}/scripts/source_to_md/{pdf,doc,excel,ppt,web}_to_md.py <文件 / URL>`。只有主题无素材 → 先主题调研再回来。

### Step 2 · 选风格 floor + 读依据
默认 business-blue。读 `design-language.md`（设计依据）+ `资产地图.md`（素材在哪），锁画布 / 配色 / 字体 / 栅格。

### Step 3 · 内容拆页型
把大纲映射到 content_type（封面 / 目录 / 章节 / 对比 / 三段… / 一段文字 / 表格 / 数据 / 结尾）：**按内容体量挑容量匹配的页型**（4 段 → seg-4，不硬塞 seg-3）。

### Step 4 · 逐页生成（调库组装 + 内容回流）
逐页：① 查该页型在 `页型配方.md` 的配方（母题 / 骨架 / 用哪件素材）；② 从 `catalog.json` 取素材，**builder / `place()` 原生注入组装**（手机 / 车 / 圆簇 / 卡 / 球… 全调库，**禁内联**）；③ 真实内容回流（标题 / 引导 / 正文 / 数据 / 图片用时填，**字号按刻度、正文别断词**）。引擎不可表达的（渐变文字 / 弧形字 / 真 3D）→ ④ 烘焙 PNG 或 ③ clone（`design-language.md` §四）。

### Step 5 · 自检
对照 `design-language.md` §五 自检清单逐条过 + §1.3 字号刻度核对；任一 NO 回 Step 4 重做（**不机械 patch**）。

### Step 6 · 渲染核对 +（可选）归档
`render_slides.ps1`(COM) 渲当前页看真图核对；满意则（可选）`add_gallery_page.ps1` 插合集册 + `sync_gallery_png.ps1` 渲该页 png（**增量、不全量重渲**）。

### Step 7 · 扩库（仅当缺页型素材）
若某 content_type 库里没有现成母题 / 风格件 → 走 `蒸馏流程.md` 10 步：从模板学一页、蒸资产入 `assets/`、加 `页型配方.md` 配方。**蒸馏 ≠ 日常生成**，是给库做加法。

---

## 定稿后 · 配讲稿（交棒 ppt-script）
定稿 pptx 后，走 `ppt-script` 给每页配**口语化讲稿**（别念稿 / 凝练 / 页间过渡）、经用户审后写进备注（见 `../ppt-script/SKILL.md`）。

---

## 收尾回灌（每个 PPT 定稿后必做，闭环触发器）
定稿后**显式执行**，否则复盘会被默默跳过：
1. 把用户本轮**每一处打回**整理成案例（症状 → 正确做法），写入 `../ppt-design/references/设计反模式案例库.md`。
2. 在该案例库的**进步指标台账**记一行 `<PPT名 版本>：打回 N 处`。**N 随版本下降 = 设计力在增强。**

---

## 备注
- 引擎源自 ppt-master（裁剪），保留 LICENSE 并署名；详见仓库 README。
- **（旧）SVG 组件管线**（`components/*.svg` + `scripts/svg_to_pptx.py`）现为 **legacy 兜底**、非主线 —— 原生蒸馏已覆盖全核心页型。`components/` / `scripts/` / `templates/` **暂留勿删**（ppt-skills 无 git 兜底；物理裁撤待提交 GitHub 前由用户定，见 `资产地图.md §三`）。
- 故障排查见 `scripts/docs/troubleshooting.md`。
