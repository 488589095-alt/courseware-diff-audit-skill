# 待办补丁清单 · 条目 schema

报告 §八的每条补丁用此结构。**只有 A 类、C 类未补、硬规则 FAIL 进清单；B 类不进**
（B 是老师增量，不是 bug）。补丁是**给人/给 ppt-script-fix 看的修改建议草案**，
伪 diff 到 TODO 级即可，不要求可直接 `git apply`。本 skill **不落盘改任何生成 skill 代码**。

## 条目字段

```yaml
patch_id: PATCH-001
category: A | C | hard_rule          # B 类不出补丁
severity: high | medium | low        # high=内容缺失/中文回退等线; medium=样式偏离; low=细节
title: 一句话问题名
target_skill: gaozhong-yingyu-handout-to-courseware   # 见 target_skills_map.md
file: scripts/extract_handout.py     # 相对 skill 目录
function: extract_tables             # 不确定写「(待定位)」
# ↓ 四段式诊断（classification_framework.md）
problem: 现象——引用机械证据 + 并排截图
root_cause: 代码层根因——哪段逻辑没处理这种内容
affected_scenarios: 会触发该段代码的所有内容类型（不止这一例）
impact_after_fix: 修复可能波及的原本正常情况
pseudo_diff: |                        # TODO 级，示意改哪、怎么改
  - # 旧：只抽段落，跳过内嵌表格
  + # 新：遍历 docx body 的 CT_Tbl，把每张表抽成 content 的 table 项
evidence:                            # 可追溯指针
  - completeness.json#H-174
  - _cmp/p41_vs_c0.png
acceptance: 修复后重跑 run_audit 应满足的可验证条件
  - completeness 中 type=表格 的 miss = 0
status: todo                         # todo | handed_to_ppt-script-fix | done
```

## 严重度判定
- **high**：A 类内容缺失（答案/正文/注释/讲义表/图漏）；`字体:中文回退等线` FAIL；WPS 冻结 FAIL。
- **medium**：样式偏离（颜色/字号/字体没贴标杆）；C 类未补的结构件；页序错。
- **low**：行距非标准（WARN）、细节标签、近似可接受的样式差。

## 分组与排序
报告 §八**按 `target_skill` 分组**，组内按 severity（high→low）排。
同一根因引发的多处现象**合并成一条补丁**（如「所有中文 run 未设 ea」是一条，不要按页拆 7 条）。

## 写补丁的纪律
1. **先看完所有证据再写**，一条补丁的修复要覆盖 `affected_scenarios` 列出的所有场景，
   不做只改一例的试探性补丁（启发式逻辑改一处易破坏另一处——沿用 ppt-script-fix 经验）。
2. **Read 目标文件确认函数仍在**再写 file/function；不确定就写「(待定位)」并在 problem 里给定位线索。
3. 每条都要有 `acceptance`（可重跑 run_audit 验证），让 ppt-script-fix/人工能闭环确认。
4. 边界1 的口径问题（讲义有+标杆无）若判不准，**不要出补丁**，在报告 §一/§三标「待用户拍板」。
