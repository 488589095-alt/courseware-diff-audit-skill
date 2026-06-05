# 目标生成 skill 索引（把根因落到具体文件/函数）

补丁清单的 `target_skill` / `file` / `function` 从这里查。所有路径相对
`/Users/gaotu/product-design-space/.claude/skills/<skill>/`。**写补丁前先 Read 目标文件确认函数仍在**
（recalled 信息可能过时）。case 目录里通常还有该 skill 的**本地化产物**（build_*.py / content.json /
extract_*.py），改动既可落到 skill 本体，也可先在 case 本地脚本验证。

## 四个 skill 的统一管线（同构）
```
讲义.docx ──[extract/parse]──▶ content.json ──[build/render]──▶ 课件.pptx
                              （+ 模版.pptx 版式 + references 样式 token）
```
- **完整性 miss（A 类）** 根因多在 **extract/parse**（某类内容没抽出）或 **content.json**（抽出了没填进去）。
- **样式偏离 / 字体回退 / 溢出** 根因多在 **build/render**（字体/颜色/字号常量、ea 未显式设、autofit）或 **references 的样式 spec**。
- **C 类未补（讲义有模版无槽位）** 根因在 **build/render** 的新增版式/文本框逻辑。
- **页序/漏页** 根因在 **build/render** 的骨架/页型编排，或 content.json 的页清单。

## 各 skill 关键文件

| skill | 学科 | 讲义解析 | 渲染 | 内容/样式规格 |
|---|---|---|---|---|
| `gaozhong-yingyu-handout-to-courseware` | 高中英语 | `scripts/extract_handout.py` | `scripts/build_pptx.py` | `references/content_schema.md` `design_tokens.md` `structure_recipe.md` `template_spec.{md,json}`；`scripts/dissect_template.py`（模版拆解Gate）`extract_styles.py` |
| `yuwen-handout-to-courseware` | 高中语文 | `scripts/extract_handout.py` | `scripts/build_pptx.py`（古诗case本地 `build_gushi_pptx.py`） | `references/content_schema.md` `font_color_spec.md` `template_map.md`；`scripts/check_overflow.py`（溢出检测） |
| `gaozhong-huaxue-handout-to-courseware` | 高中化学 | `scripts/extract_handout.py`（富提取:图片导出+上下标） | `scripts/build_pptx.py` `build_pptx_bench.py` | `references/content_schema.md` `template_map.md` `online-system-issues.md` |
| `handout-to-courseware` | 小学数学 | `scripts/parse_handout.py` + `scripts/parse_teaching_design.py`（需教学设计稿） | `scripts/render_pptx.py` | `references/typography_rules.md` `activity_to_layout_rules.md` `known_template_layouts.md` `p_number_mapping.md` |

## 常见差异 → 大致落点（先验，仍需 Read 确认）

| 差异现象（来自证据） | 大致 file / function |
|---|---|
| 讲义某表格/内嵌图没进课件（completeness miss，type=表格/image_check） | `extract_handout.py` 的表格/图片提取（化学有图片导出；其余看是否抽 docx 表） |
| 讲义某【答案】没进课件（type=答案 miss，A类） | `extract_handout.py` 答案抽取 → `content.json` → `build_pptx.py` 回填该页 |
| 中文显示等线（style hard_rule FAIL） | `build_pptx.py` 设字体处：中文 run 未显式写 `a:ea` typeface（应对所有中文 run 设 ea） |
| 答案没标红 / 颜色值不对（palette_diff / hints） | `build_pptx.py` 颜色常量 或 `references/font_color_spec.md`(语文) / `design_tokens.md`(英语) |
| 文字溢出框（截图可见） | `build_pptx.py` 的 `_fit_size`/autofit；语文有 `check_overflow.py` |
| C 类内容没新增（讲义有模版无槽位） | `build_pptx.py` 新增文本框/版式分支 + `structure_recipe.md`(英语)/`template_map.md` |
| 页型/页序错（structure order_milestones） | `build_pptx.py` 骨架编排；英语见 `structure_recipe.md` |

## 与 ppt-script-fix 的交接
补丁清单产出后，**实际改脚本交 `ppt-script-fix`**（它有结构检查 + 截图 + 视觉诊断的自动化循环）或人工。
本 skill 不改任何生成 skill 的代码。把补丁条目（含 file/function/伪diff/验收标准）贴给 ppt-script-fix 即可。
