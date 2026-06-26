# _gallery · 成果展示(原生方法学一页 = 一件成品)

> 🟡 2026-06-22(重构):**单一汇总册 + 圆簇合并** —— 只保留 `business-blue-master.pptx`(唯一可编辑 pptx,**单一真相**)+ 每页一张 png。每页 png 由 `_experiments/_toolkit/sync_gallery_png.ps1` **从汇总册渲染**(永不脱钩,1600×900);新增/替换页用 `add_gallery_page.ps1` 插临时 pptx、再渲该页 png、随即删临时。p99 `text-block-2`:「大圆+周边小圆」合并蒸为**一件** `bz-ghost-cluster`(废 ghost-orb/soft-panel/accent-chip),builder 全 `place()` 调库、零内联。catalog **43 件**、合集册 **14 页**,**核心页型全覆盖 ✅**。
> 这里是**结果**(汇总册 + 每页 png),清理流程**绝不删本目录**;过程脚本/中间版/临时单页 pptx 在 `_experiments/`,**用完即清理**(见 `蒸馏流程.md §0` 收尾两步)。
> 渲染 = PowerPoint COM 真图(所见即所得)。png 直接看,汇总册可打开再编辑。**纪律:只渲当前学/改的页,绝不全量重渲;`render_truth.ps1` 已设护栏禁止对 _preview/_gallery/assets 跑。**

## 清单

| 成品 | content_type | 母题 / 看点 | 源模板页 |
|---|---|---|---|
| `closing` | 结尾/致谢 closing | **大渐变标题 + 楼群 + 点阵球**:结尾 builder `bz-closing`(原生渐变标题/光晕) + **点阵球** `bz-dot-sphere`(vector,5655轮廓) + **四栋楼群** `bz-towers`(透明,整体/单栋 srcRect) + 球**复用** `bz-sphere-glow`+虚影;标题/副标用时传 | work_inspire_128 p128 |
| `data-chart` | 数据图表 data-chart | **缺口KPI卡 + 柱图**:缺口顶卡 `bz-stat-card` + 柱图 builder `bz-bar-chart`(渐变柱/数值/类别/趋势箭头**不压数字**,字号回模板,**形状重绘非原生chart**) + 大数字KPI(18,收回卡内);数据用时传 | work_inspire_128 p117 |
| `showcase` | 图文产品 showcase | **叠层图文**:两张圆角倒影图槽错位(占位/填图) `bz-showcase` builder + 同心涡卷装饰 `bz-deco-swirl`(一件,左原样/右镜像出血) + 渐变胶囊眉标 + 特性 chip;图片/文案用时填 | work_inspire_128 p97 |
| `text-block` | 一段文字 text-block | **单文本块·特性环变体**:页标题 + **28pt 粗引导句** + 正文段(左列) + 手机槽 + 同心 ghost 环 + **5 放射特性圈**(真图标) + 底波浪;参数化 builder `bz-text-block` | work_inspire_128 p98 |
| `text-block-2` | 一段文字 text-block | **单文本块·多视觉簇变体**:标题 + 引导句 + 正文(左) + **车 `bz-car-ev` + 双手机 `bz-phone-mock` 对角错位**(占位/填图) + **ghost 圆簇 `bz-ghost-cluster`**(大圆+周边小圆整组一件)背景,**无特性圈 / 无底波浪**;同一 builder `bz-text-block`(clusters/rings/wave/kind 参数化) | work_inspire_128 p99 |
| `people-team` | 团队/人物 people | **zigzag 成员卡网格**:双幽灵环头像占位(白色人形) `bz-team-card` + 完整背景 `bz-team-bg`(径向渐变+airy波浪) + 网格 builder `bz-team-grid` + 巨大淡水印词;照片/文字用时填 | work_inspire_128 p87 |
| `table-rank` | 表格 table | **排名/数据表**:圆顶角渐变表头 + 白卡投影 + 斑马行 + 第3列加粗;参数化 builder `bz-table-rank`(「2形状+2表格」,数据用时传) | work_inspire_128 p74 |
| `timeline-snake` | 流程/时间轴 timeline | **蛇形升序时间轴**:速度模糊背景 `bz-speed-blur-bg` + 蛇形骨架 `bz-timeline-snake`(路径+7圆) + 用时叠加图标/文字 + 页脚 `bz-footer-note` | work_inspire_128 p63 |
| `logic-cycle-loop` | 逻辑闭环 logic-diagram | **闭环循环图**:中心球 + 4 节点 + 循环箭头 + 同心光晕 + 右图标步骤(整页母题 `bz-cycle-loop` 一件) | work_inspire_128 p30 |
| `seg-multi-pillar-hub` | 多段 seg-multi | **五支柱枢纽**:中心标题六边 hub + 5 图标圆放射 + 出血外框六边 | work_inspire_128 p53 |
| `seg4-hex-hub` | 四段 seg-4 | **六边蜂窝枢纽** + 4 内容卡(给内容图形分量)+ 透明图标 | work_inspire_128 p45 |
| `seg3-ring-radial` | 三段 seg-3 | **圆环放射刻度盘**枢纽 + 3 卫星沿环弧分布 + 中心赞踩 | work_inspire_128 p36 |
| `section-light` | 章节过渡 section(亮系) | 亮系满底 + 宽幅天际线带 + 巨大渐变过渡编号(藏于楼群后) | work_inspire_128 p18 |
| `seg2-venn-compare` | 两段·对比 seg-2 | 同心探边双圆 Venn + 要点沿弧均布 | work_inspire_128 p25 |

> 命名:`<content_type>-<母题>`。同页多版只留**最终采用版**。
> 这些成品用的素材在 `skills/ppt-skills/assets/business-blue/`(矢量 `vectors/`、`catalog.json`);配方在 `references/页型配方.md`(§3.1–3.15)。
