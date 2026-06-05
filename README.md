# courseware-diff-audit

课件「生成成片 vs PPT 模版 vs 标杆 PPT」三方对比检查与诊断 skill（Claude Code / Claude Agent Skill）。

把「讲义→课件」生成 skill 产出的课件，与 **PPT 模版** 和 **参考标杆 PPT** 三方对比：
**以讲义为裁判**做完整性审计，按 **A/B/C 三分类法**归因差异，产出**诊断报告 + 待办补丁清单**
（精确到目标生成 skill 的文件/函数，但**不落盘改代码**），用于指导生成 skill 迭代优化。

> 本 skill 只**诊断**，不生成课件、不改脚本。配合「讲义转课件」生成类 skill 与脚本修复类 skill 使用。

## 三分类内核（以讲义为裁判）

| 类 | 讲义 | 标杆 | 模版 | 判定 | 处置 |
|---|---|---|---|---|---|
| **A** | 有 | 有 | — | 课件必须有 | 标杆有而课件漏/错 = **生成 skill 的 bug**，出补丁 |
| **B** | 无 | 仅标杆有 | — | 老师人工精备增量（测评/干扰项/流程图/插画/翻译…） | **忽略，非 bug** |
| **C** | 有 | 有 | 无槽位 | 讲义有内容但模版无对应版式 | 按标杆样式新增；未补则出补丁 |

## 四维对比

1. **完整性审计** — 讲义逐项是否 100% 转入课件（A 类强信号）
2. **结构/页序对账** — 课件 vs 标杆的页型/页序/页数差额（用 `role_delta`，Σ==total_delta），按 5 类老师增量桶拆分
3. **样式保真度** — 字体/字号/颜色/版式 vs 模版主题+标杆；含跨学科硬伤「中文回退等线」（数据驱动，不硬编码字体）
4. **视觉逐页对比** — 渲染缩略网格 + 标杆\|课件 关键页并排

## 用法

```bash
# 0) 定位输入（课件最新版 / 模版 / 标杆 / 讲义）并人工确认
python3 scripts/locate_inputs.py --case <case目录>

# 1) 跑审计编排器，产出全部机械证据（可 --no-screenshots 跳过渲染）
python3 scripts/run_audit.py --case <case目录>
```

产出（写到 case 目录）：`completeness.json` / `structure_diff.json` / `style_diff.json` +
`_grid/`（缩略网格）+ `_cmp/`（关键页并排）+ `audit_bundle.json`。

随后 Claude **先读完所有证据再下诊断**，按 [`references/`](references/) 做 A/B/C 归因，
写 `差异分析报告.md`（8 章节）+ 按目标 skill 分组的补丁清单。

## 设计原则

- **脚本只给事实，Claude 给判断**：脚本永不说「这是 bug」，只给「讲义有 X / 课件未命中 X / 标杆第 n 页有 X」，
  归因/分类/根因/补丁全部由 Claude 按 `references/classification_framework.md` 判定。
- **看 `by_type` 与 `content_coverage`，别只看 headline coverage**：答案/正文类 miss 才是 A 类强信号。
- **口径先声明**：讲义有+标杆无的内容做不做，取决于「贴标杆」还是「讲义全覆盖」口径，判不准就标「待用户拍板」。

## 目录结构

```
SKILL.md                  入口（流程 + 纪律）
scripts/
  locate_inputs.py        定位四类输入 + 目标 skill
  run_audit.py            编排器（结构提取 → 机械差异 → 视觉渲染 → 汇总）
  completeness_audit.py   讲义↔课件机械匹配（hit/partial/miss）
  structure_reconcile.py  页型/页序/页数对账 + 5 类增量桶
  style_fidelity.py       字体/字号/色 + 硬规则（数据驱动）
  visual_compose.py       缩略网格 + 并排拼图（Pillow）
  _common.py              pptx/docx 统一解析（实时读 + dump 回退，兼容两种 docx dump 格式）
  dump_pptx.py / dissect_template.py / take_screenshots.py / structural_checks.py  复用脚本（自包含）
references/                A/B/C 框架 / JSON schema / 报告模板 / 补丁 schema / 目标 skill 索引
evals/regression_notes.md  两个 case 的回归验证记录
```

## 依赖

- Python 3：`python-pptx`、`python-docx`、`lxml`、`Pillow`
- 视觉渲染（pptx → PDF → PNG）需 **macOS + Microsoft PowerPoint + PyObjC（Quartz）**；
  缺失时 `run_audit.py` 自动跳过视觉、仅出结构/样式诊断（不阻断）。

## 安装为 Claude Skill

```bash
git clone https://github.com/488589095-alt/courseware-diff-audit-skill.git
cp -r courseware-diff-audit-skill ~/.claude/skills/courseware-diff-audit   # 文件夹名需与 skill 名一致
```
