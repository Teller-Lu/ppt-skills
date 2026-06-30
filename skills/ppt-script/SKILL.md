---
name: ppt-script
description: >
  Use when adding speaker notes / 讲稿 / 备注 to a FINISHED deck — 给做好的 PPT 配口语化讲稿、写进备注。
  Reads the approved 逐页设计稿 + the built slides → per-page spoken script → writes into pptx notes.
  Runs AFTER ppt-skills has produced the pptx. Not for designing slides (ppt-design) or building them (ppt-skills).
---

# PPT Script · 讲稿（写进备注）

> 做 PPT 的**最后一道工序**：给定稿 pptx 配**口语化讲稿**，写进幻灯片备注。
> 轻引擎、重火候——和 `ppt-design` 同构：核心是《讲稿铁律》，写备注的机制是现成的（python-pptx notes API）。
> **本 skill 跑在 `ppt-skills` 把 pptx 做完之后**；不设计页面（那是 `ppt-design`）、不做页面（那是 `ppt-skills`）。

**核心流水线**：`读《设计稿》+ 成品页 → 逐页四件套(承上/核心/展开/钩子) → 讲稿.md + 软时长报告 → 交用户审 → 写进 pptx 备注`

---

> [!CAUTION]
> ## 🚨 前置门 + 交接协议
>
> 1. **前置门**：开工先确认有「**定稿 pptx** + 它的《逐页设计稿》」。缺定稿 pptx → 引导先走 `ppt-skills` 做完再来。
> 2. **先审后写**：产物是可审的 `讲稿.md`；**经用户审批**才写进 pptx 备注（同设计稿的"先审后用"）。
> 3. **别念稿**（最高铁律，详见《讲稿铁律》）：讲稿绝不逐字复述屏幕上的字——屏幕给"看的要点"，讲稿给"说的展开"。
> 4. **真实**：只讲页面真实素材支持的内容，不编数据 / 结论。

---

## 知识（生成必读）

| 资源 | 用途 |
|---|---|
| **`references/讲稿铁律.md`** ⭐ | 口语化 / 凝练 / 别念稿 / 页间过渡 的七条铁律。逐页对照，必读 |
| `references/讲稿模板.md` | `讲稿.md` 的格式 + 写备注用的 `@@@SLIDE@@@` payload 约定 |
| `references/个人讲稿风格.md` | 【预留 · v1 未实做】有真实讲稿样本时怎么蒸口头风格 |

---

## 主流程

### 第 1 步 · 读两头
- **《逐页设计稿》**：每页 ①一句话核心、②为什么、备注"与邻页关系"，全局 3 条主线、目标总时长。**讲稿主料。**
- **成品页文本**：用 python-pptx 从定稿 pptx 提每页文字（或看渲染 png），**只为"别念稿"校验**——知道屏上有哪些字、好补充不复读。

### 第 2 步 · 逐页四件套
每页写：
1. **承上过渡句**——把上一页连过来 + 锚定属哪条主线（首页 = 开场白：今天讲什么 / 为什么值得听）。
2. **口语核心句**——设计稿①核心翻成大白话。
3. **展开**——背景 / 为什么 / 举例 / 强调，**补充而非复读** bullet。
4. **下页钩子**（可选）——引出下一页。

全程对照《讲稿铁律》——尤其"别念稿""口语化""页间过渡"。

### 第 3 步 · 汇成讲稿.md + 软时长报告
用 `references/讲稿模板.md` 汇成 `讲稿.md`。按 **~200 字/分**估：每页标"约 X 秒"，文末报"全篇预计 Y 分钟 vs 目标 Z 分钟"，**超了只提示**（不自动砍），由用户定夺。

### 第 4 步 · 交用户审
`讲稿.md` 交用户审批、改。**批准后**才写备注。

### 第 5 步 · 写进 pptx 备注
把每页**口语正文**（承上 + 核心 + 展开 + 钩子，**不含 ⏱ 行**）按页序拼成备注源 txt，每页之间单独一行 `@@@SLIDE@@@`，**段数 = 幻灯片数**。跑：
```
python skills/ppt-script/scripts/write_notes.py <定稿.pptx> <备注源.txt> -o <带讲稿.pptx>
```
默认输出 `<deck>.noted.pptx`（不覆盖定稿）。

---

## 备注
- 写备注的机制复用引擎能力（python-pptx notes API），本 skill 只管讲稿的**内容与火候**。
- 个人风格：v1 通用铁律打底；将来给真实讲稿样本，按 `references/个人讲稿风格.md` 蒸口头风格覆盖在通用规则之上。
