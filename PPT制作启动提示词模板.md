# PPT 制作 · 启动提示词模板

> **用途**：每要做一个新 PPT，**复制下面 §二 代码块**，把所有 `【填:…】` 换成本次 PPT 的信息，**粘进一个新对话的第一条消息**即可——新对话瞬间定位任务/能力/路径/铁律，直接从《逐页大纲》开干。
> **为什么单开对话**：每个 PPT 背景任务不同，单开互不污染；本模板把"固定能力 + 铁律 + 流程"钉死，你只填"这次做什么"。
> **维护**：ppt-skills 素材库/路径有变时，回来更新 §二 的【你的能力】段即可。

---

## 一、填写指引（对应 §二 里的 `【填:…】`）

| 占位 | 填什么 | 常见选项 / 示例 |
|---|---|---|
| PPT 名称 / 版本 | 这份 PPT 叫什么 + 版本 | 四驱扭矩寻优·立项汇报 v1.1 |
| 一句话目标 | 给谁看、要达成什么 | 给领导/专家汇报，认可价值+可实施，通过立项 |
| 主文档路径 | 内容唯一来源(.md/.docx) | `D:\...\项目介绍.md` |
| 资料清单路径 | 已填的《资料采集清单》(有就给) | 同目录 `项目PPT资料采集清单.md`；无则写"约束见下" |
| 类型 / 听众 / 场合 | — | 立项 / 评审 / 进展 / 结题 / 学术；领导 / 专家 / 客户 / 内部 |
| 时长 / 页数 | — | 10min/~10页；15min/~15页… |
| 主线 | 听众最该记住的 1–3 条 | ①价值 ②怎么做 ③难点 |
| 风格 | geely / business-blue / 混合 | geely 封面结尾 + business-blue 浅色中间页 |
| 数据 / 真实图 | 有→路径；无→暂无 | 真实图见项目介绍.md |
| 其它硬要求 | logo / 模板 / 保密 / 截止… | 以 geely-master.pptx 为基调 |
| 输出路径 | 成品 pptx 存哪 | `项目目录\xxx_v1.1.pptx` |

> **最少填**：主文档路径 + 类型/听众 + 时长/页数 + 主线 + 风格，即可启动。

---

## 二、提示词模板（复制此代码块 → 填好 `【填:…】` → 粘进新对话首条消息）

```text
你是本《【填:PPT 名称】（【填:版本，如 v1.1】）》的【制作方】。目标：【填:一句话——给谁看、要达成什么】。全程中文；不要提任何费用/计费。

【先读这几样(按序)】
1. 主文档(全部内容来源)：【填:主文档绝对路径】
2. 已填资料清单(约束/风格/主线/真实图指向，若有)：【填:资料清单路径，或写"无，约束见下"】
3. 本工作区根手册 CLAUDE.md(已自动加载) + 记忆里所有 ppt-* 条目(设计/蒸馏/分工/避坑经验，务必看)

【你的能力 = ppt-skills plugin 个人模板库(读文件用，不是装好的插件)，含两个 skill】根目录 D:\Lu.Yao7\VehicleTellerLu\Agent\Shared\skills-creator\ppt-skills
- ★设计前脑 ppt-design(做 PPT 第一步、先用它出设计稿)：skills\ppt-design\SKILL.md + references\(表达式词典.md 内容→最优表达→红线、逐页设计稿模板.md 四问、设计反模式案例库.md 避坑、范本-立项v1.1.md 样板)
- 制作 skill 运行手册：skills\ppt-skills\SKILL.md
- ★美化主脑(制作时必读)：skills\ppt-skills\references\design-language.md(景深/光影/占版/字号刻度/出件四选一/自检清单/避坑；表达决策已前置到 ppt-design)
- 页型配方：skills\ppt-skills\references\页型配方.md
- 资产地图：skills\ppt-skills\references\资产地图.md(素材在哪、怎么 place 调用)
- geely 公司风：skills\ppt-skills\assets\geely\(builder gy-signature.py = 纯白封面/丝绸封面/THANKS 结尾/中间页角标；catalog.json；rasters\gy-cover-silk-bg.jpg)
- business-blue 商务蓝：skills\ppt-skills\assets\business-blue\catalog.json(封面/目录/章节/对比/三段/四段/多段/逻辑闭环/时间轴/表格/数据图表/团队/图文/一段文字/多图/甘特 bz-gantt-grid/架构树 bz-arch-tree/架构背景 bz-arch-bg/结尾…) + builders/vectors/rasters + _palette.pptx(素材带便签，肉眼挑件)
- 样张(先翻一遍看效果)：_gallery\*.png 与合集册 business-blue-master.pptx / geely-master.pptx
- 引擎(python-pptx 管线)：D:\Lu.Yao7\github\ppt-master
- 工具脚本：skills\ppt-skills\_experiments\_toolkit\(render_slides.ps1 逐页渲染核对、add_gallery_page.ps1、sync_gallery_png.ps1)

【本次任务约束】
- 类型 / 听众 / 场合：【填:如 立项汇报；领导+评审专家】
- 时长 / 目标页数：【填:如 10 分钟 / ~10 页】
- 主线(听众最该记住的几条，按此铺)：【填:如 ①价值 ②怎么做 ③技术难点】
- 风格：【填:如 geely 公司风为主(封面+结尾)+中间页 business-blue 浅色素材；浅色；带单位 logo】
- 数据：【填:有→路径/说明；无→暂无】　真实图：【填:见某文档/文件夹；或暂无】
- 其它硬要求：【填:如 以 geely-master.pptx 为模板基调；保密标记；截止日期…】

【铁律(踩了就返工)】
- **表达优先、别硬套模板**(详见 ppt-design 的《表达式词典》)：每页先定"用什么表达"——**结构/流程/甘特/层级 优先 mermaid(我写 mermaid 代码 → 你用 Obsidian Mermaid Exporter 导成图 → 我当真实图嵌入)或用户真实图**；框图慎用、手搓矢量是最后手段；**模板是地板不是紧身衣，不贴就别套**(改用表格/大图/纯排版)
- 配图真实性：真实结果/数据/图必须真实(取自上面主文档/素材)，严禁编造、AI 生成、假数据凑；缺真实素材的页→放占位并标【需提供】请用户补，绝不假图凑数
- 文字一律微软雅黑，且烤进每个 run 的 ea/cs(否则注入回退宋体)；字号/占版对齐模板，别 AI 小字 + 胆怯留白
- 真实 logo 不进公开素材库：模板留占位框；实战取真 logo 于 skills\ppt-skills\_experiments\company-style\icons_tmp\ 填进去
- DLP 水印(© 方能演亦 / EagleCloudWatermark)严禁出现在成品
- 操作 PowerPoint 用 COM 时：开演示前先记已开的 Presentations 数，只 Close 自己开的、绝不 Quit(用户可能正开着别的)；.ps1 脚本纯 ASCII
- 改/扩 skill 只改 GitHub 源(Shared\skills-creator)，不碰 .claude 插件目录；work-pc 上 git push 必须用工作区根 push_via_proxy.sh

【开工步骤】
第 1 步（设计，走 ppt-design）：读完上面输入 + 翻 _gallery 样张，用 ppt-design 逐页**四问**产出《逐页设计稿》给我审(别直接做 pptx)。格式见 skills\ppt-design\references\逐页设计稿模板.md，照 范本-立项v1.1.md 的填法。每页一块：
  〔页N 标题〕角色(封面/过渡/内容/结尾) | ① 一句话(这页唯一核心) | ② 表达式(查《表达式词典》指到一类:mermaid/真实图/库件id/表格/纯排版) | ③ 为什么(对红线自证,禁"模板里有") | ④ 真实素材(缺的标【需提供】)
  页数按上面约束，主线按上面顺序；封面/结尾用 geely。
第 2 步（制作，走 ppt-skills）：我批准设计稿后，你按稿用 builder/place 调库做成可编辑 pptx(输出到：【填:输出路径，如 项目目录\xxx_vN.pptx】)。设计稿的"表达式/真实素材"即制作依据，别另起设计。
第 3 步：用 render_slides.ps1 逐页渲染 PNG，自己对照 design-language 自检(字号/无重叠/真素材非占位/水印净) → 改 → 给我核对 → 定稿。
第 4 步（收尾回灌）：定稿后把我每处打回写进 skills\ppt-design\references\设计反模式案例库.md，并在进步指标台账记一行"<PPT名 版本>：打回 N 处"。

现在开始第 1 步：读输入，走 ppt-design 四问，产出《逐页设计稿》。
```

---

## 三、备注

- 新对话与本工作区共享 `CLAUDE.md` + 记忆，模板里再把关键点钉死 = 双保险。
- **第一交付物锁死为《逐页设计稿》**(走 ppt-design 四问，不直接做 pptx)——先审主线/表达/分页，避免一口气做歪。
- 已存档的实例可放本目录或各项目目录(如四驱扭矩寻优立项 v1.1 的填好版)。
