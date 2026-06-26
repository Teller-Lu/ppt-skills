# ppt-skills · 更稳定接近成品的个人风格 PPT skill

> 在 ppt-master「文档 → 可编辑 PPTX」引擎之上，用人工打磨的精美模板当"质量地板"，把真实内容适配进去，产出**更稳定、更接近成品**的可编辑 PPTX。
> A personal-style, template-floored, *representation-first* PPT skill for Claude Code — turns documents into editable, on-brand PPTX.

---

## 这是什么

AI 做 PPT 最大的短板是**即兴设计带来的质量方差**。本 skill 做两件事压住它：

1. **质量地板**——默认设计从"AI 即兴"换成"人类已验证的模板"，从根上抬高下限；
2. **表达优先**——在动手做之前，先强制"想清楚每页要表达什么核心"，避免一味套模板 / 堆美化。

成品是**可编辑 pptx**（原生 DrawingML，文字 / 配色 / 素材都能在 PowerPoint 里继续改），不是一张导出图。

## 两个 skill：先设计，后制作

做 PPT 分两步，各由一个 skill 负责，中间以一份**《逐页设计稿》**交接：

| skill | 职责 |
|---|---|
| **`ppt-design`**（设计前脑） | 做 PPT 第一步。逐页**四问**——这页一句话核心 / 用什么表达 / 为什么 / 需要哪些真实素材——产出《逐页设计稿》，经你审批。只设计，不碰素材库。 |
| **`ppt-skills`**（制作） | 拿到批准的设计稿，用原生素材库 + builder 组装成可编辑 pptx，渲染核对、自检。入口有"前置设计门"，无批准设计稿不开工。 |

**表达优先**意味着：结构 / 流程 / 甘特 / 层级优先用 mermaid 或真实图，而不是手搓框图；模板是质量地板、不是枷锁，不贴就换更简单的表达（表格 / 大图 / 纯排版）。每次成稿后把被打回的点回灌到"设计反模式案例库"，**越用越准**。

## 风格与覆盖

- **business-blue 商务蓝**：原生素材库覆盖全部核心页型——封面 / 目录 / 章节 / 对比 / 多段枢纽 / 逻辑环 / 时间轴 / 表格 / 数据图 / 团队 / 图文 / 甘特 / 架构树 / 结尾。
- **geely 公司风**：纯白封面 / 丝绸封面 / THANKS 结尾 / 中间页角标。
- 样张见 [`_gallery/`](_gallery/)（合集册 + 每页 png）。

## 安装（Claude Code 插件）

本仓库是一个 Claude Code plugin。在 Claude Code 里：

```
/plugin marketplace add Teller-Lu/ppt-skills
/plugin install ppt-skills@ppt-skills
```

依赖：Python 3.10+，以及 ppt-master 引擎（见 `skills/ppt-skills/requirements.txt` 与下方致谢）。

## 目录结构

```
skills/
  ppt-design/    设计前脑：SKILL.md + references(表达式词典 / 逐页设计稿模板 / 反模式案例库 / 范本)
  ppt-skills/    制作：SKILL.md + references(design-language 等) + assets(business-blue / geely)
_gallery/        样张：合集册 + 每页 png
docs/            方法论 / 规格文档
```

## 当前进度

- ✅ business-blue 核心页型原生覆盖（封面 → 结尾）
- ✅ geely 公司风（封面 / 结尾 / 中间页）
- ✅ 双 skill 流程：`ppt-design` 设计 → 审批 → `ppt-skills` 制作 → 回灌
- 🔜 同类型多风格变体；更多公司风模板；data-chart 升级为真·可编辑图表

## 致谢与许可

- 本项目以 **MIT** 许可发布，见 [`LICENSE`](LICENSE)。
- 工程引擎源自 **ppt-master**（MIT，© 2025-2026 Hugo He，裁剪集成）；其版权与许可声明保留于 [`LICENSE`](LICENSE) 末尾的第三方声明。
- 模板素材为作者个人蒸馏、打磨，供个人风格生成参考。
