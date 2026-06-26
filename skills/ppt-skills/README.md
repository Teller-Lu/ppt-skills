# skills/ppt-skills/ — skill 本体说明

> 这是被 Claude 发现并加载的**真正的 skill**(`skills/<name>/SKILL.md` 是发现入口)。
> 项目全局结构见 [../../docs/文件系统说明.md](../../docs/文件系统说明.md)。

## 这层里有什么

| 路径 | 是什么 | 谁产出 |
|---|---|---|
| `SKILL.md` | **编排大脑**:触发词 + 9 条执行纪律 + 8 步流程。Claude 加载它来知道"怎么做 PPT" | 我们写 |
| `requirements.txt` | Python 依赖(裁剪版:无生图/配音) | 我们写 |
| `scripts/` | **引擎**(89 个 py,从 ppt-master 裁剪复制)。核心是 `svg_to_pptx`(SVG→可编辑 pptx)、`pptx_to_svg`(解析模板)、`svg_quality_checker`(技术门)、`finalize_svg`、`svg_editor`(预览)、`source_to_md`(文档转换) | 复制+裁剪 |
| `references/` | **规则知识**(25 md):`executor-base` / `strategist` / `shared-standards`(引擎安全构造的黑白名单)/ `canvas-formats` / 图文版式等。Executor 手写 SVG 前要读 | 来自 ppt-master |
| `styles/` | **风格库(质量地板)**。每套风格一个文件夹,含 `brand.md`(审美 DNA)+ `gate.calibration.json`(美学门阈值)+ `page-inventory.md`。目前只有 `academic-defense`(从你 25 页提炼) | 我们建设 |
| `components/` | **组件库**:可复用的整页模板。`catalog.json` 是索引,`ad-*.svg` 是组件本体(引擎安全 + 占位内容),`*.assets/` 是烘焙资产 | 我们建设 |

## 三层心智
- **大脑** = `SKILL.md`(指挥)· **肌肉** = `scripts/`(干活)· **知识** = `references/`(规则)· **审美** = `styles/` + `components/`(你的风格)。

## 现状(2026-06-13)
- 引擎双向跑通;`academic-defense` 风格 + 9 个组件已建;**但组件较朴素,正在"深化学习 + 加华丽"阶段**。
- 详细进度见 [../../docs/实施计划.md](../../docs/实施计划.md) §进度。
