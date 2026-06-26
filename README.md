# ppt-skills · 更好用的个人风格 PPT skill

> 🟡 2026-06-23:**多图素材回炉 + 蒸馏改分工**——多图 p103-106 母题入库反复出错(过度拆分 / 抽错几何 / 造重复)后,**改定新惯例:图片+矢量母题由用户从原模板手动复制进 `_palette.pptx`,我只贴便签 + 登 catalog + 渲 _preview**(见记忆 `ppt-distill-division-of-labor`)。用户手排 `_palette`(**40 页**:30 基线 + 10 多图,glow/frame/wave 成对并页);我整理便签、剥全册 `©方能演亦` DLP 水印(50 run)、提升为正式版(旧版备份在 `_experiments/_toolkit/`)。catalog **对齐 palette = 57 件**(去 `bz-tag-bar` 重复等 6 件之 catalog 条目、补 `bz-phone-frame` 真机框;`bz-fpg-bg`=`bz-tilt-bg` 同图已 dedup);`_preview` **38 张**(图片页不渲)。教训:**"蒸馏素材到库"重在 vectors/rasters/catalog + palette/_preview;视觉抽取交用户、我管元数据**。
>
> 2026-06-22(重构):**文档按读者拆三档 + 圆簇合并**——① `design-language.md` 瘦身为**生成必读主脑**(原则/字号/出件/自检/坑),拆出 `references/页型配方.md`(15 页型 cookbook,按需查一节)+ `references/蒸馏流程.md`(学一页/扩库 10 步);② **重写 `SKILL.md` 走原生蒸馏主线**(SVG 管线降为 legacy);③ 修字号矛盾→统一刻度(内容页标题/引导 28、正文≈14、签名页按模板放大,§1.3);④ p99「大圆+周边小圆」合并蒸为**一件** `bz-ghost-cluster`(废 ghost-orb/soft-panel/accent-chip)。catalog **43**、合集册 **14 页**、palette **30**。改前已备份(`_backups/`)。**🎉 business-blue 全核心页型(封面→结尾)原生覆盖完成。**
>
> 在 **ppt-master 引擎**之上,用**我大量人工打磨的精美模板**当"质量地板",做一个**更稳定接近成品**的个人风格 PPT skill。
> 本目录 = 规划与开发工作区;skill 成品在 `skills/ppt-skills/`。归属 `Shared/skills-creator/`。
> 语言:统一中文(claude 回复 + 内容生成)。

## 这是什么
把默认设计从"AI 即兴"换成"人类已验证模板",**从根上抬高质量下限**。借 ppt-master 的「文档 → 可编辑 PPTX」工程引擎,把它最大的短板——**靠 AI 即兴设计导致的质量方差**——用模板补上。个性化是副产品,**"更稳定接近成品"才是目的**。

## 模板三型(学习素材的形态 · 详见 `docs/素材模型与组件库规格.md`)
1. **单风格型** — 一套完整 PPT、全篇统一风格 → 喂 `styles/`(整体一致性"地板")。例:学术答辩 25 页。
2. **单内容汇总型** — 单个内容类型的多种风格(如"目录汇总""数据图表汇总")→ 喂零件库。
3. **汇总型** — 全 content_type × 多风格,素材量最大 → 最大零件库。**当前在学:`work_inspire_128`(128 页商务蓝汇总册)。**

## 两套方法(并行,待收敛)
| 方法 | 载体 | 状态 |
|---|---|---|
| (旧)**SVG 组件管线** | `components/*.svg` + `scripts/svg_to_pptx.py` | **legacy 兜底**(暂留勿删,待提交 GitHub 前裁定) |
| (新·主线)**原生 DrawingML 蒸馏** ⭐ | `design-language`+`页型配方` + `assets/business-blue/` | **全核心页型已覆盖 ✅**,`SKILL.md` 主流程 |

> **待决**(提交 GitHub 前):原生覆盖完成后,是否裁掉 SVG 管线 / `components/`。见 `references/资产地图.md §三`。

## 🗺 当前学习路线(原生重做 · 每 content_type 一页 · 广度优先)
源:`work_inspire_128`(汇总型);每页走 `references/蒸馏流程.md` 10 步,蒸馏入 `assets/business-blue/` + `references/页型配方.md` 配方。代表页由 `_experiments/work_inspire_128/page-inventory.md` 选定。

| content_type                  | 代表页                            | 原生状态              |
| ----------------------------- | ------------------------------ | ----------------- |
| 封面 cover                      | p8                             | ✅                 |
| 目录 toc                        | p17                            | ✅                 |
| 章节过渡 section                  | p18 亮 / p19 暗                  | ✅                 |
| 两段·对比 seg-2                   | p25                            | ✅                 |
| **三段 seg-3**                  | **p36 圆环放射枢纽** | ✅ 模板级         |
| 四段 seg-4                      | p45 六边蜂窝枢纽                    | ✅ 模板级         |
| 多段 seg-multi                  | p53 五支柱枢纽                     | ✅ 模板级         |
| 逻辑 / 环图 logic-diagram         | p30 闭环(左cycle+右图标步骤)      | ✅ 模板级         |
| 流程 / 时间轴 timeline             | **p63 蛇形升序(7 阶段)**         | ✅ 模板级         |
| 表格 table                      | p74 排名表(圆顶角表头+斑马)        | ✅ 模板级         |
| 数据图表 data-chart               | p117 缺口KPI卡+柱图              | ✅ 模板级         |
| 团队 / 人物 people                | p87 zigzag 成员卡                | ✅ 模板级         |
| 图文产品 showcase                 | p97 叠层图+特性chip              | ✅ 模板级         |
| 一段文字 text-block               | p98/p99 单文本块+配图           | ✅ 模板级(加宽)  |
| 结尾 closing                    | p128 渐变标题+楼群+点阵球        | ✅ 模板级         |
| 多图排版 multi-image            | p103-106(手机簇/三屏/中心/倾斜) | ✅ 全学(加宽批)   |
| **甘特 gantt**                  | **p109 阶梯甘特(网格+任务条)**  | ✅ 模板级         |
| **逻辑图 logic-tree**           | **p111 平台架构树(5列层级)**    | ✅ 模板级         |
| 组织架构 / 荣誉 / 伙伴 / 联系       | 待挑                            | ⬜ 后补              |

> 同类型多风格变体(如三段 p31/p36、数据饼 / 环 / 折线)**广度铺满后再回头加**。
> 📋 **页级学习台账**(哪些模板页已学 / 未学 / 本批队列,128 页不遗漏)见 [`_experiments/work_inspire_128/学习台账.md`](_experiments/work_inspire_128/学习台账.md)。

## 后续阶段
- **Phase A(当前)** — business-blue 全 content_type **原生覆盖**(上表)。
- **Phase B** — 再挑几套**单风格型**好模板做新 floor(如用新方法**重做学术风**;旧方法效果一般)。
- **Phase C** — 从**单内容汇总型**补"同类型多风格"变体。

## 目录导航
| 路径 | 内容 |
|---|---|
| `skills/ppt-design/` 🟢 | **设计前脑**(2026-06-26 新增·做 PPT 第一步):`SKILL.md`(四问主流程) + `references/`(表达式词典 / 逐页设计稿模板 / 设计反模式案例库 / 范本-立项v1.1);出《逐页设计稿》经审批再交 ppt-skills 制作 |
| `skills/ppt-skills/SKILL.md` | skill 运行手册(流水线 + 全局纪律;入口前置设计门 + 收尾回灌) |
| `skills/ppt-skills/references/design-language.md` | ⭐ **生成必读·设计主脑**:原则 + 字号刻度 + 出件四选一 + 自检清单(§五)+ 避坑(§六) |
| `skills/ppt-skills/references/页型配方.md` | **页型 cookbook**:15 种 content_type 层叠配方(做某页型查一节) |
| `skills/ppt-skills/references/蒸馏流程.md` | **学一页 / 扩库流程**(10 步蒸馏标准;仅扩库时读) |
| `skills/ppt-skills/references/资产地图.md` | 素材在哪 / 怎么调用 + `templates/icons` 图标库 + components(legacy)说明 |
| `skills/ppt-skills/assets/business-blue/` | **原生素材库**:catalog / vectors / rasters / **builders**(参数化件如表格) / _palette / _preview |
| `skills/ppt-skills/assets/geely/` | 🟢 **GEELY 公司风**(2026-06-23 新增·与 business-blue 平级):`builders/gy-signature.py`(一件 4 版式:纯白封面/丝绸封面/THANKS/中间页角标) + `rasters/gy-cover-silk-bg.jpg`(丝绸背景) + `catalog.json`;真实 logo 不入库、留占位框,实战 logo 在 `_experiments/company-style/icons_tmp/` |
| `_gallery/`(索引 `_gallery/README.md`) | **成果**:合集册 `business-blue-master.pptx`(唯一 pptx·单一真相) + 每页 png(从汇总册派生);新增 / 换页用 `_toolkit/add_gallery_page.ps1` + `sync_gallery_png.ps1`(增量) |
| `docs/`(索引 `docs/README.md`) | 规划 / 方法论 / 规格文档(三型模型 / 内容适配 / 美学门 …) |
| `_experiments/work_inspire_128/` | 模板 128 页真渲染 + `page-inventory.md`(页型↔代表页) |
| `_inbox/` | 原始 `.pptx` 模板投放处 |

## 状态
- 🟢 **2026-06-26 双 skill 化**:新增 `ppt-design` 设计前脑(四问主流程 + 表达式词典 + 设计反模式案例库 v0.1 + 范本),制作 skill `ppt-skills` 入口加**前置设计门**、收尾加**回灌案例库 + 打回 N 进步指标**;`design-language` 的表达决策门**迁出**到 ppt-design(留指针);改名 `ppt-make` **延后**(见 `docs/ppt-design_双skill拆分方案.md` 附录 A)。
- 日期:2026-06-18 · 本机 Python 3.10 ✓ · 引擎源 `D:\Lu.Yao7\github\ppt-master`
- 原生已学:**封面 / 目录 / 章节 / 对比 / 三段 / 四段 / 多段 / 逻辑闭环 / 时间轴 / 表格 / 团队人物 / 图文产品 / 数据图表 / 结尾致谢**(14 类型,**核心页型全覆盖 ✅**);`catalog` **60 件**(对齐用户手排 `_palette` **43 页** / `_preview` **41 张**,含 `bz-gantt-grid`(甘特)+`bz-arch-tree`(架构树)+`bz-arch-bg`(架构页背景);多图 builder `bz-multi-image` 一件 4 版式;一段文字 p98/p99:`bz-text-block`+`bz-wave-layers`+`bz-ghost-circles`+**`bz-ghost-cluster`**[大圆+周边小圆整组一件]+`bz-phone-mock`+`bz-car-ev`;closing 蒸楼群 / 点阵球 / 结尾,球复用 `bz-sphere-glow`;**全库文字已烤微软雅黑** §6.23;蒸骨架别蒸内容 §6.24;**禁 builder 内联画、只调库** 蒸馏流程 §0 步4铁律);文档三档:`design-language`(主脑)+ `页型配方`(配方)+ `蒸馏流程`(学一页);成品 `_gallery/`(合集册 **20 页**(+多图 4 + 甘特/架构树 2) + 每页 png)。
- 流程:每页学成→builder 建临时 pptx→`add_gallery_page.ps1` 插汇总册→`sync_gallery_png.ps1` 渲该页 png(**不留单独 pptx**)、清 `_experiments`(**护住 _gallery/合集/_toolkit**,§0 步 9/10);**资产增量只动当前页、绝不全量重渲**;每页开工前 /compact + 重申全程中文。
- 下一步:**核心页型已全覆盖**(封面→结尾);加宽批 **一段文字 p98+p99 已学**(`text-block` 特性环 + `text-block-2` 多视觉簇,同一 builder)。多图排版 p103-106 **4 页全学完**(`bz-multi-image` 一 builder 4 版式,合集册 14-17 页);**队列剩:甘特图 p107-109 / 逻辑图 p110-114 各学 1 页**(用户暂停,转实战 PPT 任务)。其后 —— ① 同类型多风格变体 ② 提交 GitHub 前裁定 SVG 管线(`资产地图.md §三`) ③ data-chart 改真·可编辑图表(待用户定)。
