# Academic Defense · 24 页 → 组件分类清单

> 解析 24 页逐页归类(content_type),作为组件提炼(阶段 4)的基础,并"生长"出新页型。
> 信号取自 `_experiments/academic_defense_25/_analysis.json`。

## 逐页分类

| 页 | content_type | 判定信号 |
|---|---|---|
| 01 | `cover` | "PPT灵感手册"96px + 副标题 + 答辩人/导师/专业 |
| 02 | `toc` | "目录 CONTENTS" + 01–05 条目 |
| 03 | `section` | 巨型 "01" 458px + "PART ONE" + 章节标题 |
| 04 | `concept` ✦新 | "基本概念" + 定义 + 中心词"都市农业"+ 多要素卡 |
| 05 | `seg-3` | "研究方法设计":实证调查/定量分析/文献分析法(polygon 箭头) |
| 06 | `seg-multi` | KPI(184/200/175份)+ 文献理论/初步分析/现场访谈 |
| 07 | `comparison` | "国内 vs 国际"两栏 + 地图图形(path×184) |
| 08 | `comparison` | "保税备货 vs 一般贸易"对比表(5 行,line×12) |
| 09 | `process-timeline` | 1994-2000…2018-Present 时间分期 |
| 10 | `process-timeline` | "研究流程"样本→检测→处理→分析(polygon 箭头) |
| 11 | `logic-diagram` | "产业生态优势"产业/创新/科研/尖峰 簇状关系 |
| 12 | `data-chart` | "[chart]" + 销售额/增速 + 数据来源/单位 |
| 13 | `seg-4` | 装饰大数字 "6"(Prism 221px)+ 4 个优势条目 |
| 14 | `table` | 4 列(诱导元素/设施/视认距离/作用)× 多行 |
| 15 | `logic-diagram` | "全域治理"全域水系/网格化/无死角(ellipse×9) |
| 16 | `logic-diagram` | "数据治理需求"需求传导/数据类型关系 |
| 17 | `seg-multi` | "传统试验问题"5 个问题卡 + 理论+物理试验 |
| 18 | `content` | "研究综述"方法+不足之处 |
| 19 | `seg-3` | "配电网规划"运行方式/网络结构/灵活性 三段 |
| 20 | `process-timeline` | "数字化布局"投入→营销→运营→生态 层级 |
| 21 | `image-layout` / `screenshot` ✦ | image×8(8 张图网格)+ 标题 + 3 要点 |
| 22 | `summary` ✦新 | 巨型 01/02/03(Bebas 184px)+ "总结" 3 点 |
| 23 | `references` ✦新 | "参考文献" [1]–[6] 引文(tspan×101) |
| 24 | `closing` | "THANKS" 318px + "敬请指正"(庞门正道粗书体) |

## 频次(哪些组件最值钱)

- 关系/分段类最多:`logic-diagram`×3、`process-timeline`×3、`seg-3/4/multi`×6 → 学术 deck 的主力。
- 框架齐全:cover / toc / section / summary / references / closing 各 1。
- 数据/表格:`table`×2、`comparison`×2、`data-chart`×1。
- `image-layout`×1(8 图网格)。

## 生长出的新页型(并入全局清单)

- `concept` 概念/定义页(中心词 + 多要素)
- `summary` 总结页(大数字 + 要点)
- `references` 参考文献页
- (模板库还有 `pyramid` 金字塔,见 `E:\PPT模版\金字塔.pptx`)

## 给阶段 4(组件提炼)的备注

- 每页 = 一个整页组件(粒度按约定)。
- 提炼时剔除装饰水印 "方能演亦"。
- `color_roles` 映射用 brand.md 的语义角色表;`slots` 按上表"判定信号"里的可变内容(标题/条目数/KPI 数值/行列…)。
