---
name: courseware-diff-audit
description: 课件「生成成片 vs PPT模版 vs 标杆PPT」三方对比检查与诊断 skill。以讲义为裁判做完整性审计，按 A/B/C 三分类法归因差异，产出诊断报告 + 按目标 skill 分组的待办补丁清单（不落盘改代码）。覆盖四维：完整性审计 / 结构页序对账 / 样式保真度 / 视觉逐页对比。当用户说"对比课件和标杆"、"课件差异分析"、"检查课件完整性"、"课件和模版/标杆差在哪"、"为什么和标杆不一样"、"出差异分析/页面绑定/补丁清单"、"课件审计/质检"时使用。注意：本 skill 只诊断不生成课件（生成用 handout-to-courseware / yuwen-handout-to-courseware / gaozhong-yingyu-handout-to-courseware / gaozhong-huaxue-handout-to-courseware），只诊断不改脚本（改脚本用 ppt-script-fix）。
---

# 课件对比检查（差异诊断 + 补丁清单）

把「生成的课件 PPT」与「PPT 模版」「标杆 PPT」三方对比，**以讲义为裁判**，
找出差异、归因、产出诊断报告 + 待办补丁清单，指导生成 skill 迭代。

**核心原则**：脚本只产生**机械事实**；A/B/C 分类、根因、补丁全部由你按 references 判定。
**唯一落盘报告** = case 目录下的 `差异分析报告.md`；**绝不改其他 skill 代码、绝不改课件 pptx**。

## 三分类内核（详见 references/classification_framework.md）
- **A**：讲义有 + 标杆有 → 课件必做；标杆有而课件漏/错 = **生成 skill 的 bug**，出补丁。
- **B**：讲义无 + 仅标杆有（测评/干扰项/流程图/插画/翻译…）= 老师精备增量，**忽略，非 bug**。
- **C**：讲义有 + 模版无槽位 → 应按标杆样式新增；未补则出补丁。

## 流程

### 0. 定位输入（先确认再跑）
```bash
python3 {SKILL_DIR}/scripts/locate_inputs.py --case <case目录>
```
扫出 课件(最新vN+全版本)/模版/标杆/讲义 与目标生成 skill，写 `audit_inputs.json`。
**先把摘要给用户确认**（多版本时确认对比哪一版；缺标杆/讲义会降级，要说明）。

### 1. 跑审计编排器（产生全部证据）
```bash
python3 {SKILL_DIR}/scripts/run_audit.py --case <case目录>
# 仅结构/样式、跳过渲染（快）：加 --no-screenshots
# 指定课件版本：--courseware <某版本.pptx>
```
依次做：结构提取(dump) → 机械差异(完整性/结构/样式) → 视觉渲染(缩略网格+关键页并排)，
汇总到 `audit_bundle.json` 并打印「请 Read 这些证据」。

产出证据（写到 case 目录）：
| 文件 | 内容 | 维度 |
|---|---|---|
| `completeness.json` | 讲义逐项 → 课件 hit/partial/miss + 命中页 | ① 完整性（A 类强信号） |
| `structure_diff.json` | role_delta / 页数差额 / 5类增量桶 / 页面绑定 | ② 结构页序 |
| `style_diff.json` | 硬规则(回退等线/WPS/行距) + 调色板差 + 校准线索 | ③ 样式保真 |
| `_grid/ _cmp/` | 缩略网格 + 标杆\|课件 关键页并排 | ④ 视觉逐页 |
| `_dump_课件_audit.txt` | 课件结构 dump（人读） | — |

> 截图需 macOS + PowerPoint（AppleScript 导 PDF → Quartz 渲 PNG）。标杆只有 dump 时无标杆截图（仅课件网格）；标杆有 pdf 时自动渲染。失败会降级继续，不阻断结构/样式诊断。

### 2. Claude 归因（核心，先看完证据再下判断）
**先 Read 完三份 JSON + `_grid/` 缩略网格 + 全部 `_cmp/` 并排图**，再动笔（沿用 ppt-script-fix「看完所有证据再诊断」纪律）。逐条把机械事实判成 A/B/C：
- 判定流程、4 个边界口径（**贴标杆 vs 讲义全覆盖**、图片目检、role_delta 算账、partial≠缺失）→ `references/classification_framework.md`
- 怎么读三份 JSON 的每个字段 → `references/diff_schema.md`
- A/C 类写**四段式诊断**：问题 / 根因 / 受影响场景 / 修改后影响。

### 3. 写报告 + 补丁清单（唯一落盘）
- 按 `references/report_template.md` 写 `差异分析报告.md`（8 章节：完整性审计 / R2忽略 / R3新增 / 页数差额 / 样式校准 / 视觉对比 / 根因 / 补丁清单）。
- A 类 + C 类未补 + 硬规则 FAIL → 按 `references/patch_item_schema.md` 转补丁条目，
  用 `references/target_skills_map.md` 落到目标 skill 的**具体文件/函数**，按 skill 分组。
- 实际改脚本**交 ppt-script-fix 或人工**——把补丁条目贴过去即可。

## 关键纪律
- **脚本给事实，你给判断**：脚本永不说「这是 bug」，只说「讲义有 X / 课件未命中 X / 标杆第 n 页有 X」。
- **看 by_type 与 content_coverage，别只看 headline coverage**：答案/正文类 miss 才是 A 类强信号；
  短标题弱匹配、留白表格的 partial 不是缺失。
- **页数差额用 role_delta**（Σ==total_delta），不用相似度绑定（会被共享语篇误绑）。
- **口径先声明**：讲义有+标杆无的内容做不做，取决于本 case 是「贴标杆」还是「讲义全覆盖」口径，
  判不准就标「待用户拍板」，别硬判成 bug。
- **中文回退等线** = 跨学科硬伤（中文 run 未显式设 `a:ea`），直接出补丁。

## 脚本清单
| 脚本 | 职责 | 来源 |
|---|---|---|
| `locate_inputs.py` | 定位四类输入 + 目标 skill | 新写 |
| `run_audit.py` | 编排器（仿 ppt-script-fix/test_driver.py） | 新写 |
| `completeness_audit.py` | 讲义↔课件机械匹配 | 新写 |
| `structure_reconcile.py` | 页型/页序/页数对账 + 增量桶 | 新写 |
| `style_fidelity.py` | 字体/字号/色 + 硬规则（数据驱动，不硬编码字体） | 新写 |
| `visual_compose.py` | 缩略网格 + 并排拼图（Pillow） | 新写 |
| `_common.py` | pptx/docx 统一解析（实时读 + dump 回退，含两种 docx dump 格式） | 新写 |
| `dump_pptx.py` `dissect_template.py` | 结构 dump / 模版拆解 | 拷贝自 gaozhong-yingyu，同源 |
| `take_screenshots.py` `structural_checks.py` | 截图 / WPS·行距硬规则 | 拷贝自 ppt-script-fix，同源 |

> 拷贝的 4 个脚本与源 skill 同源、自包含；源 skill 更新时手动同步。
