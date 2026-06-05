# 机械差异 JSON schema 说明（怎么读三份证据）

三个审计脚本各产一份 JSON（run_audit 写到 case 目录）。脚本只给事实，**不下结论**。

---

## completeness.json —— 完整性审计（以讲义为裁判，A 类强信号）

```jsonc
{
  "handout": "_dump_讲义.txt", "handout_source": "dump|live",
  "courseware": "第9讲_课件_v8.pptx", "courseware_source": "live",
  "items": [
    { "id": "H-012", "handout_ref": "P9", "section": "固化情节一：坚持克服",
      "type": "答案|知识标签|解析|注释|章节标题|正文|表格|留白",
      "text_excerpt": "...", "is_image": false, "priority": "high|medium|low|notes|skip",
      "match": { "status": "hit|partial|miss|image_check|skip",
                 "courseware_pages": [9], "similarity": 0.97 } }
  ],
  "summary": {
    "total_graded": 121, "hit": 108, "partial": 13, "miss": 0, "image_check": 0,
    "coverage": 0.946,            // (hit + 0.5·partial) / total
    "content_coverage": 0.96,     // 仅 答案/注释/正文/表格 —— 比 coverage 更能反映内容转入
    "by_type": { "答案": {"hit":26}, "解析": {"hit":8,"miss":15}, ... },
    "skipped_blanks": 5
  },
  "miss_items": [ {id, ref, type, section, excerpt} ]   // miss + image_check 汇总，逐条判 A/B/C
}
```
怎么用：
- **看 `by_type`，不是只看 coverage**。答案/注释/正文 的 miss = A 类强信号；解析的 miss 看口径（边界1）。
- `image_check` 项必须看并排截图目检（边界2）。
- `partial` 多为短标题弱匹配 / 留白表格，不等于缺失（边界4）。
- 切块规则与 priority 含义见 completeness_audit.py 顶部 docstring。

---

## structure_diff.json —— 结构/页序/页数对账（课件 vs 标杆）

```jsonc
{
  "courseware": {"file","total","source"}, "benchmark": {"file","total","source"},
  "role_histogram": { "courseware": {"正文/其它":62,...}, "benchmark": {"选择题":22,...} },
  "page_count": {
    "total_delta": 29,                 // 标杆total − 课件total
    "role_delta": [                    // ★主信号：每页唯一角色，Σdelta == total_delta
      {"role":"选择题","benchmark":22,"courseware":4,"delta":18}, ... ],
    "benchmark_more_pages": 32, "courseware_more_pages": -3,
    "benchmark_only_by_binding": 18,   // 仅供配图，会偏小（共享语篇被误绑）—— 别拿来算账
    "unbucketed_benchmark_pages": 5, "warn_unbucketed": false, "note": "..."
  },
  "benchmark_only_buckets": [ {"bucket":"互动题+干扰项","pages":[12,13,...],"count":10} ],
  "benchmark_only_pages": [ {"benchmark_page","role","bucket","summary"} ],
  "courseware_only_pages": [ {"courseware_page","role","summary","best_benchmark_sim"} ], // C类新增候选
  "page_binding": [ {"benchmark_page","benchmark_summary","role","courseware_page","similarity"} ],
  "order_milestones": { "courseware":[{name,page}], "benchmark":[...] }
}
```
怎么用：
- 页数差额叙事**用 `role_delta`**（边界3）。标杆多出的角色 → 逐条对到 5 类增量桶（B 类）。
- `benchmark_only_buckets` 是初分，`warn_unbucketed` 为真就看截图细分。
- `courseware_only_pages` = 课件有标杆无 → 多为 C 类新增（封面/分隔/题源…），核对是否讲义内容。
- `page_binding` 给 visual_compose 配并排图；低 similarity 的对子要重点看（样式漂移）。

---

## style_diff.json —— 样式保真度（课件 vs 模版+标杆，数据驱动）

```jsonc
{
  "courseware": "...", "benchmark": "...",
  "baseline": { "source":"template_spec.json|none", "major_font":{...}, "minor_font":{...}, "theme_colors":{...} },
  "hard_rules": [ {"rule":"字体:中文回退等线","status":"FAIL","pages":[...],"detail":"..."}, ... ],
  "palette": {
    "courseware": {"colors":[{value,count}], "fonts_ea":[...], "fonts_latin":[...], "fonts_name":[...], "sizes":[...]},
    "benchmark":  {...}            // 标杆来自 dump 时为 shape 级（每形状取首 run），采样偏少
  },
  "palette_diff": {
    "colors_benchmark_only": ["FF5050","A076CE"],  // 标杆用、课件没用 = 样式没贴上
    "colors_courseware_only": [...], "sizes_benchmark_only": [...],
    "fonts_benchmark_only": [...], "fonts_courseware_only": [...]
  },
  "style_calibration_hints": [ {"element":"强调红(答案/标签)","issue":"标杆用红而课件未见红色字",...} ]
}
```
怎么用：
- `hard_rules` 的 FAIL = 跨学科硬伤（WPS 冻结 / 中文回退等线），直接出补丁。WARN（行距/未显式设ea）视情况。
- `palette_diff.colors_benchmark_only` + `style_calibration_hints` = 颜色/字号/字体没对齐标杆的线索。
- **注意非对称**：课件实时读=run 级（最全），标杆 dump=shape 级（偏少）→ `*_courseware_only` 偏多属正常，
  重点看 `*_benchmark_only`（标杆有而课件无更可信）。有标杆 live pptx 时两边都 run 级，更对称。
- `style_calibration_table`（题/诗/答案/分隔标题等**语义元素**逐行）需你用本 palette + 截图填，脚本不臆造语义分组。
