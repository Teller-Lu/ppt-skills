# ppt-master 解析蓝本

> 本文是改造的**底图**。要安全地改一个系统,先得懂它**每个设计为什么这么做**——所以本文重点不只是"是什么",更是"为什么"。
> 一手资料:ppt-master 自带 `docs/zh/technical-design.md`(WHY 级技术设计)+ `skills/ppt-master/SKILL.md`(编排权威),2026-06-12 通读。

---

## 0. 一句话定性 + 设计哲学

**ppt-master =「源文档 → AI 手写 SVG → 工程化转成可编辑 PPTX」的多角色串行流水线。**

它的设计哲学有两句话最关键,直接划定了能力边界:

1. **「AI 是你的设计师,不是完工师」** —— 产出的 pptx 是一份**设计稿**,不是成品;要精良,需要你在 PowerPoint 里做"精装修"。目标是消除 90% 从零开始的工作量,而非最后一公里。
2. **「工具的上限是你的上限」** —— 它放大你已有的设计品味与判断,**不能替你产生品味**。

→ 这两句既是它的诚实,也是它最大的**实践短板**(见 [优化方向与改进设计.md](优化方向与改进设计.md))。

---

## 1. 系统总览(主流水线)

```
用户输入 (PDF/DOCX/XLSX/URL/Markdown/纯文本/主题)
   │
   ├─① 源内容转换  source_to_md/*  → 归一化为 Markdown(Strategist 的事实源)
   ├─② 创建项目    project_manager.py init <名> --format <格式>
   ├─③ 模板处理    默认跳过(自由设计);仅当用户给显式模板路径才启用
   ├─④ Strategist  八项确认(⛔阻塞)→ design_spec.md + spec_lock.md
   ├─⑤ 图片获取    (条件)AI 生图 / 网络搜图        ← 我们裁掉
   ├─⑥ Executor    逐页手写 SVG → svg_output/ ; 质量门 ; 讲稿 notes/total.md
   │                 (实时预览 server.py --live @ localhost:5050)
   └─⑦ 后处理导出  total_md_split → finalize_svg → svg_to_pptx
                     └→ exports/*.pptx(原生形状版,主交付)+ backup/
```

**三阶段心智**:① 内容理解与设计规划(Strategist)→ ② AI 视觉生成(Executor 手写 SVG,产物是**设计稿**)→ ③ 工程化转换(脚本把 SVG 逐元素翻成 DrawingML 原生对象,可点选、可改色)。

---

## 2. 角色系统:三个角色,但是**单主代理切换**,不是子代理(回答"包含哪些 agent")

| 角色 | 运行模式 | 进场先读 | 产出 |
|---|---|---|---|
| **Strategist 策略师** | 与用户**协商**(开放、对话、可回退) | `references/strategist.md` | `design_spec.md` + `spec_lock.md` |
| **Image_Generator 图片师**(条件触发,**我们裁掉**) | 产出生图 prompt / 搜图 | `image-base.md` + `image-generator.md` | 图片资源 + 状态 |
| **Executor 执行师** | **严格产出 XML**(不准即兴、不准漏属性) | `executor-base.md` + `shared-standards.md` + 一个风格文件 | 逐页 SVG |

**关键事实:这三个是"同一个主代理在切换角色",不是并行子代理。** ppt-master 给了三条互相支撑的理由:

1. **为什么不用并行子代理**:页面设计依赖**完整上游上下文**(Strategist 的配色、图片是否获取成功、前几页的视觉节奏)。子代理只能拿到上下文的过期局部快照,产出的 deck 会**逐页视觉漂移**。同一逻辑也**禁止分批生成**(如一次 5 页)——分批加速上下文压缩,一致性掉得比省的快。
2. **为什么角色专属 reference,而不是一个超大 prompt**:Strategist 跑"协商"模式(可回退),Executor 跑"严格 XML"模式(不准即兴)。塞进同一个 prompt 等于逼模型在一轮里持守互相矛盾的纪律——会犯所有混合 prompt 的病。拆开后每个角色只加载自己要的。
3. **角色切换协议**:切角色前必须 `read_file references/<role>.md`,既把新鲜指令载入上下文覆盖前一模式的漂移,又在对话里留下可见的审计标记。

---

## 3. 七步流水线(逐步 + 前置 GATE / 阻塞点)

| 步 | 名称 | 关键动作 | 阻塞? |
|---|---|---|---|
| 1 | 源内容处理 | 非 Markdown 一律先转 MD;只给主题→先跑 `topic-research` | 否 |
| 2 | 项目初始化 | `project_manager.py init`;源文件 `import-sources --move` 进 `sources/` | 否 |
| 3 | 模板选项 | **默认自由设计**;只有用户给**显式目录路径**(含 `kind: brand/layout/deck`)才启用 | 否 |
| 4 | **Strategist** | 读 strategist.md → **八项确认(⛔阻塞,唯一硬停点)** → 出 design_spec + spec_lock | ⛔ 是 |
| 5 | 图片获取 | 条件触发(资源表有 ai/web 行);全 user/formula 则跳过 | 否 |
| 6 | **Executor** | 起实时预览 → 逐页手写 SVG → **质量门**(0 error)→ 讲稿 | 否 |
| 7 | 后处理导出 | `total_md_split` → `finalize_svg` → `svg_to_pptx`(三条命令**逐条**跑) | 否 |

**八项确认**(Step 4 的唯一阻塞 gate,打包呈现、一次确认):① 画布格式 ② 页数范围 ③ 受众 ④ 风格目标 ⑤ 配色 ⑥ 图标用法 ⑦ 排版(含公式渲染策略)⑧ 图像用法。
打包且单一的理由:设计选项**互相关联**(配色影响图标、影响排版),一起决才一致;分散确认会引入互相矛盾的输入,被迫回退重做。**确认后,流水线一路自动跑到结束,不再有中断点。**

---

## 4. 九条全局执行纪律(为什么这么"官僚")

存在的根因一句话:**LLM 的默认行为是"让我这一轮把整个问题搞定",而这恰恰是串行流水线最不该有的形状。** 九条各关闭一种**真实复现过**的失败模式:

| # | 规则 | 关闭的失败模式 |
|---|---|---|
| 1 | 串行执行 | 乱序 |
| 2 | BLOCKING = 硬停 | AI 代用户做设计决策 |
| 3 | 禁跨阶段打包 | 阶段混淆 |
| 4 | 进入前过 GATE | 前置条件未满足 |
| 5 | 禁投机执行 | 提前给后续步骤"备料"(如策划阶段就写 SVG) |
| 6 | SVG 生成禁子代理 | 子代理上下文丢失 |
| 7 | 只逐页顺序生成 | 分批漂移 |
| 8 | 每页前重读 spec_lock | 长 deck 色彩/字体漂移 |
| 9 | **SVG 必须手写,禁脚本批量生成** | 跨页视觉一致性丢失 |

> **第 9 条尤其重要,且和我们的改造直接相关**:脚本批量生成 SVG 的路线**试过、在一个分支上被放弃了**——因为跨页视觉一致性依赖"带完整上游上下文的逐页手写",生成器脚本复现不了。
> 注意区分:**让主代理在完整上下文里"组装"预先备好的组件 ≠ 被禁的"生成器脚本"**。这条边界是我们做"组件化复用"(见优化文档)时必须守住的红线。

---

## 5. 核心枢纽:为什么是 SVG

SVG 是整条流水线的中枢,是逐一排除其他方案后的胜者:

| 候选 | 为什么被排除 |
|---|---|
| 直接生成 DrawingML | 极繁琐(一个圆角矩形几十行嵌套 XML),AI 训练数据少,质量不稳 |
| HTML/CSS | "文档流"世界观 vs PPT"画布"世界观,根本冲突;`<table>` 变不成独立形状 |
| WMF/EMF | 微软自家矢量格式,但 AI 几乎没有训练数据 |
| SVG 当图片嵌入 | 丢失全部可编辑性,和截图没区别 |

**SVG 胜出**:它和 DrawingML **同一世界观**(绝对坐标二维矢量),`<path>↔<custGeom>`、`<rect rx>↔prstGeom roundRect`、`linearGradient↔gradFill` 一一对应——转换是"两种方言间的精确翻译"。且唯一同时满足:**AI 能可靠生成、人能在浏览器直接预览调试、脚本能精确转换**。
**viewBox 用像素不用 EMU**:像素空间让 AI 排版和人类调试都无歧义,EMU 换算只在导出时发生一次。

---

## 6. 抗漂移三件套(质量稳定性的核心机制)

| 产物 | 给谁 | 内容 |
|---|---|---|
| `design_spec.md` | 人读 | 设计的**"为什么"**(受众、风格目标、配色理由、页面大纲) |
| `spec_lock.md` | 机器读 | Executor 必须**字面照搬**的**"是什么"**(HEX 色值、确切字体名、图标库、图片资源表) |
| 逐页重读 spec_lock | 机制 | SKILL.md 强制"每页生成前 `read_file spec_lock.md`",让数值在 20+ 页里字面一致 |

`update_spec.py` 把生成后的改动传播到所有 SVG,但**故意只支持 `colors.*` 和 `typography.font_family`**——其他字段(字号/图标/图片)需要语义级理解,风险不值得自动批处理,手改后重做受影响页。

---

## 7. 模板系统:brand / layout / deck(我们要"反转"的地方)

| 种类         | 目录                        | 含什么                    | frontmatter    |
| ---------- | ------------------------- | ---------------------- | -------------- |
| **brand**  | `templates/brands/<id>/`  | 仅身份:配色/字体/logo/语气/图标风格 | `kind: brand`  |
| **layout** | `templates/layouts/<id>/` | 仅结构:画布/页结构/页型/SVG 名册   | `kind: layout` |
| **deck**   | `templates/decks/<id>/`   | 全复刻:身份 + 结构 + 中段       | `kind: deck`   |

**默认自由设计,模板 opt-in,且只在用户给出显式路径时触发**(无名称模糊匹配、不主动推荐、不基于内容自动套)。ppt-master 给的理由:

> **"模板是地板,但很容易变成天花板"** —— 它会把整个 deck 锁进模板自有的视觉惯用语,无视内容本身想怎样被呈现。自由设计让版式跟着内容走,而不是套一层固定语法。

补充不对称:**布局是 opt-in,图表和图标不是**——布局才是"锁定视觉惯用语"的那层(地板/天花板),图表图标是不施加 deck 级风格约束的复用原语。

→ **这正是我们项目的核心改造点**:把"模板=可选天花板"反转为"模板=默认质量地板"(详见优化文档 §三)。

---

## 8. 图片体系(我们裁掉,但要懂它的思路)

- **AI 生图三维系统**:`rendering`(视觉风格家族)× `palette`(deck 的 HEX 在图里怎么用)× `type`(单图内部构图),Strategist 阶段就**锁进 spec_lock**,之后每张图的 prompt 从锁定的 rendering+palette 组装——这是 spec_lock 抗漂移思想"往像素上游推一层"。
- **图文版式词表**:`Primary 主结构`(骨架)+ `Modifier 修饰层`(装饰),72 条编号技法**自由组合**。它**显式鼓励复合**,专门对抗 AI 的失败模式——不是"叠太多",而是"用太少"(每页堆成裸卡片网格 → 视觉扁平的"AI 默认感")。

> 即便我们不生图,**"风格三维锁定"和"版式复合对抗扁平"这两个思想,对我们做风格库极有借鉴价值**。

---

## 9. 质量门:只校验**技术**,不校验**美学**(我们要补强的地方)

`svg_quality_checker.py` 在 `svg_output/` 上跑,**前置于后处理**(后处理会重写 SVG 掩盖违规)。它检查:

- 禁用 SVG 特性(黑名单:mask/style/@font-face/foreignObject/animate*/script…)
- viewBox 不符、spec_lock 漂移
- **图表几何**(柱高/扇角/刻度位置——LLM 映射数据→像素常有 10–50px 误差)

**严重性模型**:`error` 阻塞(Executor 带设计意图**重写**该页)、`warning` 不阻塞、**故意无 auto-fix**(机械 patch 会丢失设计意图,交付更难看的页)。

> **关键缺口**:它**不检查**对齐、留白、层级、对比度、溢出这些**美学**维度。这类只有 `visual-review`(opt-in、AI 主观 rubric)才碰,且**非默认**。
> → 我们要加一道**可量化的"美学质量门"**(见优化文档 §二.2)。

---

## 10. 后处理与导出(四产物 + 转换器)

| 产物 | 服务的工作流 | 不可替代性 |
|---|---|---|
| `svg_output/` | 唯一**手写**源、手工编辑入口、质检 | 流水线里唯一非派生目录 |
| `svg_final/` | IDE / 浏览器即时预览 | pptx 在 IDE 打不开;svg_output 图标图片是外链渲染不全 |
| `exports/*.pptx`(native) | **主交付** | 唯一能在 PowerPoint 原生改尺寸/改色/改样式的产物 |
| `exports/*_svg.pptx`(快照,需 `--svg-snapshot`) | 跨平台单文件分发 | 自包含多页;各家 Office 都能开 |
| `backup/<ts>/svg_output/` | 不重跑 LLM 从冻结源重建 pptx | Executor 原始 SVG 唯一留存 |

- `svg_finalize/` 有**两种消费者**:写盘(给 svg_final)+ 内存(native 转换器直接调,用于展开图标 `<use>`、扁平化多行 `<tspan>`)。
- **native 转换器逐元素派发**:每种形状一个窄翻译器,可单独调试单测;一页质量 = 各局部转换之和。
- **Office 兼容模式默认开**:为旧版 PowerPoint(<2019)生成 PNG 兜底,与原生形状并存。

---

## 11. 独立工作流(6+ 个,稀疏触发,都 opt-in)

`topic-research` / `template-fill` / `create-template` / `create-brand` / `verify-charts` / `customize-animations` / `live-preview` / `visual-review` / `resume-execute`。
**都不进默认流水线**——否则给多数用户跑无意义步骤(增加延迟和失败面)。每个 `workflows/<name>.md` 自包含、按需加载,所以 prompt 开销也是 opt-in。

---

## 12. 文件地图(大脑 / 肌肉 / 知识 / 素材)

| 层 | 位置 | 角色 |
|---|---|---|
| 大脑(编排) | `SKILL.md` | 7 步 + 9 纪律,指挥全程 |
| 知识(规则) | `references/`(strategist / executor-* / shared-standards / canvas-formats / image-layout-*) | 角色定义 + 技术规范 |
| 肌肉(工具) | `scripts/`(~90 py) | 源转换、建项目、SVG 质检、finalize、**svg_to_pptx(核心)**、pptx_to_svg、实时预览… |
| 模板(可选) | `templates/`(brands/layouts/decks/charts/icons) | opt-in 风格件 |
| 样例(只读) | `examples/` | 成品画廊,流水线**从不读取** |

---

## 13. 对本项目的启示(承接优化文档)

- **保留(命脉)**:`svg_to_pptx`(SVG→可编辑 pptx)、`finalize_svg`、质量门框架、`spec_lock` 抗漂移机制、角色切换、SKILL.md 编排骨架、`pptx_to_svg`(吃模板)。
- **裁掉**:生图(`image_backends` 13 家)、配音(edge-tts)、生图三维体系、动画(可选先不要)。
- **改造杠杆**(→ [优化方向与改进设计.md](优化方向与改进设计.md)):
  1. 模板从"opt-in 天花板"**反转**为"默认质量地板";
  2. 在技术质量门之上**补一道可量化的美学质量门**;
  3. **组件化复用**(守住第 9 条红线:主代理带上下文组装,不是生成器脚本);
  4. 数据图表走**真实绘图/数据绑定**,消灭几何误差类返工。
