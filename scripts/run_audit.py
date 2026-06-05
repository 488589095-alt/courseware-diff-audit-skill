# -*- coding: utf-8 -*-
"""
run_audit.py — 课件对比检查编排器（仿 ppt-script-fix/test_driver.py 四段式）。

流程：定位输入 → 结构提取(dump) → 机械差异(完整性/结构/样式) → 视觉渲染对比
     → 汇总 audit_bundle.json + 打印「请 Read 这些证据再做 A/B/C 归因」。

脚本只产生**事实证据**；A/B/C 分类、根因、补丁清单由 Claude 按 references 完成，
最终写 差异分析报告.md（唯一落盘报告，不由本脚本生成）。

用法:
  python3 run_audit.py --case <dir>                 # 全量
  python3 run_audit.py --case <dir> --no-screenshots  # 跳过渲染（快，仅结构/样式）
  python3 run_audit.py --case <dir> --courseware <指定版本.pptx>
"""
import argparse
import json
import sys
import traceback
from pathlib import Path

SKILL_DIR = Path(__file__).parent
sys.path.insert(0, str(SKILL_DIR))

import locate_inputs
import completeness_audit
import structure_reconcile
import style_fidelity
import visual_compose


def _ensure_dump(src_pptx_or_docx, kind, out_txt):
    """src 为实时文件时生成 dump（供 Claude 阅读）；已存在 dump 直接用。"""
    import dump_pptx
    out_txt = Path(out_txt)
    try:
        if kind == "pptx":
            dump_pptx.dump_pptx(str(src_pptx_or_docx), str(out_txt))
        else:
            dump_pptx.dump_docx(str(src_pptx_or_docx), str(out_txt))
        return str(out_txt)
    except Exception as e:
        print(f"  △ dump 失败 ({kind}): {e}")
        return None


def _screenshot(pptx, out_dir):
    import take_screenshots
    return take_screenshots.run(str(pptx), str(out_dir))


def _pdf_screens(pdf, out_dir):
    import take_screenshots
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    return take_screenshots._pdf_to_png(str(pdf), str(out_dir))


def select_pairs(structure_diff, max_pairs=12):
    """从 page_binding 选有代表性的 标杆↔课件 配对（按标杆页序均匀采样 + 含低相似度页）。"""
    if not structure_diff:
        return []
    bound = [b for b in structure_diff["page_binding"] if b["courseware_page"]]
    if not bound:
        return []
    # 低相似度(0.55~0.72)的绑定优先看（样式漂移嫌疑）
    drift = [b for b in bound if b["similarity"] < 0.72][:4]
    # 其余按标杆页序均匀采样
    step = max(1, len(bound) // max_pairs)
    sampled = bound[::step]
    chosen, seen = [], set()
    for b in drift + sampled:
        key = (b["benchmark_page"], b["courseware_page"])
        if key not in seen:
            seen.add(key)
            chosen.append([b["benchmark_page"], b["courseware_page"]])
        if len(chosen) >= max_pairs:
            break
    return chosen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True)
    ap.add_argument("--courseware", default=None, help="指定课件版本（默认最新 vN）")
    ap.add_argument("--no-screenshots", action="store_true")
    a = ap.parse_args()

    case = Path(a.case).resolve()
    print(f"\n{'#'*64}\n# 课件对比检查: {case.name}\n{'#'*64}")

    # ── 阶段0 定位输入 ──
    inp = locate_inputs.locate(str(case))
    locate_inputs.print_summary(inp)
    (case / "audit_inputs.json").write_text(json.dumps(inp, ensure_ascii=False, indent=2),
                                            encoding="utf-8")

    cw_pptx = a.courseware or inp["courseware"]["latest"]
    if not cw_pptx:
        sys.exit("❌ 无课件可对比")
    cw_pptx = Path(cw_pptx)
    handout = inp["handout"]["live"] or inp["handout"]["dump"]
    bench = inp["benchmark"]["live"] or inp["benchmark"]["dump"]
    bench_live = inp["benchmark"]["live"]
    bench_pdf = inp["benchmark"]["pdf"]
    tspec = inp["template"]["spec_json"]
    if not tspec and inp["template"]["live"]:
        try:
            import dissect_template  # noqa
            import subprocess
            subprocess.run([sys.executable, str(SKILL_DIR / "dissect_template.py"),
                            inp["template"]["live"], "-o", str(case)], timeout=120)
            if (case / "template_spec.json").exists():
                tspec = str(case / "template_spec.json")
        except Exception as e:
            print(f"  △ 模版拆解失败: {e}")

    bundle = {"case": str(case), "case_name": case.name, "inputs": inp,
              "courseware_used": str(cw_pptx), "artifacts": {}, "errors": []}

    # ── 阶段1 结构提取（dump，供 Claude 阅读）──
    print(f"\n{'─'*60}\n[1/4] 结构提取 dump")
    bundle["artifacts"]["dump_courseware"] = _ensure_dump(cw_pptx, "pptx", case / "_dump_课件_audit.txt")

    # ── 阶段2 机械差异 ──
    print(f"\n{'─'*60}\n[2/4] 机械差异")
    if handout:
        try:
            comp = completeness_audit.run(handout, str(cw_pptx))
            (case / "completeness.json").write_text(json.dumps(comp, ensure_ascii=False, indent=1),
                                                    encoding="utf-8")
            bundle["artifacts"]["completeness"] = str(case / "completeness.json")
            s = comp["summary"]
            print(f"  完整性: coverage {s['coverage']} (内容类 {s['content_coverage']}), "
                  f"{s['miss']} miss + {s['image_check']} 待目检")
        except Exception as e:
            bundle["errors"].append(f"completeness: {e}")
            print(f"  ✗ 完整性失败: {e}")
    else:
        print("  △ 无讲义，跳过完整性（无法做 A/B/C 归因）")

    structure_diff = None
    if bench:
        try:
            structure_diff = structure_reconcile.run(str(cw_pptx), bench)
            (case / "structure_diff.json").write_text(json.dumps(structure_diff, ensure_ascii=False, indent=1),
                                                      encoding="utf-8")
            bundle["artifacts"]["structure_diff"] = str(case / "structure_diff.json")
            pc = structure_diff["page_count"]
            print(f"  结构: 差额 {pc['total_delta']} (标杆多 {pc['benchmark_more_pages']}/课件多 {pc['courseware_more_pages']})")
        except Exception as e:
            bundle["errors"].append(f"structure: {e}")
            print(f"  ✗ 结构失败: {e}\n{traceback.format_exc()}")
    else:
        print("  △ 无标杆，跳过结构对账")

    try:
        style = style_fidelity.run(str(cw_pptx), bench, tspec)
        (case / "style_diff.json").write_text(json.dumps(style, ensure_ascii=False, indent=1),
                                              encoding="utf-8")
        bundle["artifacts"]["style_diff"] = str(case / "style_diff.json")
        fails = [h for h in style["hard_rules"] if h["status"] == "FAIL"]
        print(f"  样式: 硬规则 {len(fails)} FAIL; " +
              (f"硬伤: {fails[0]['rule']}" if fails else "无硬伤"))
    except Exception as e:
        bundle["errors"].append(f"style: {e}")
        print(f"  ✗ 样式失败: {e}")

    # ── 阶段3 视觉渲染对比 ──
    print(f"\n{'─'*60}\n[3/4] 视觉渲染")
    cw_png = case / "_shots" / "课件"
    bm_png = case / "_shots" / "标杆"
    grids_pairs = None
    if a.no_screenshots:
        print("  跳过渲染 (--no-screenshots)")
    else:
        try:
            _screenshot(cw_pptx, cw_png)
            bundle["artifacts"]["screenshots_courseware"] = str(cw_png)
        except Exception as e:
            bundle["errors"].append(f"screenshot_cw: {e}")
            print(f"  ✗ 课件截图失败(需PowerPoint): {e}")
        # 标杆截图：优先 live pptx，其次 pdf
        try:
            if bench_live:
                _screenshot(bench_live, bm_png)
                bundle["artifacts"]["screenshots_benchmark"] = str(bm_png)
            elif bench_pdf:
                _pdf_screens(bench_pdf, bm_png)
                bundle["artifacts"]["screenshots_benchmark"] = str(bm_png)
            else:
                print("  △ 标杆无 pptx/pdf，仅 dump → 无标杆截图（无法并排，仅课件网格）")
        except Exception as e:
            bundle["errors"].append(f"screenshot_bm: {e}")
            print(f"  ✗ 标杆截图失败: {e}")
        # 拼图
        try:
            pairs = select_pairs(structure_diff)
            vis = visual_compose.run(cw_png if cw_png.exists() else None,
                                     bm_png if bm_png.exists() else None,
                                     case, pairs)
            bundle["artifacts"]["visual"] = vis
            grids_pairs = vis
        except Exception as e:
            bundle["errors"].append(f"visual: {e}")
            print(f"  ✗ 拼图失败: {e}")

    # ── 阶段4 汇总 + 指引 ──
    (case / "audit_bundle.json").write_text(json.dumps(bundle, ensure_ascii=False, indent=1),
                                            encoding="utf-8")
    print(f"\n{'═'*64}\n证据已就绪 → audit_bundle.json")
    print("═"*64)
    print("请 Claude 依次 Read 以下证据，再按 references 做 A/B/C 归因与补丁清单：")
    for k, v in bundle["artifacts"].items():
        if isinstance(v, dict):
            for kk, vv in v.items():
                if isinstance(vv, list):
                    for p in vv:
                        print(f"  · {p}")
                elif vv:
                    print(f"  · {vv}")
        elif v:
            print(f"  · {v}")
    if bundle["errors"]:
        print("\n⚠️ 过程告警:")
        for e in bundle["errors"]:
            print(f"  - {e}")
    print(f"\n下一步：按 references/report_template.md 写 差异分析报告.md，"
          f"A/C 类差异按 references/patch_item_schema.md 生成补丁清单。")


if __name__ == "__main__":
    main()
