# -*- coding: utf-8 -*-
"""
locate_inputs.py — 扫 case 目录，定位对比检查的四类输入并推断目标生成 skill。

四类输入（实时文件优先，dump 回退）：
  · 课件   = 生成成片 pptx（取最新 vN，但列全部版本供 Claude 选）
  · 模版   = template_spec.json（优先）/ _dump_模[板版].txt / 实时模版 pptx
  · 标杆   = _dump_标杆*.txt / 实时标杆 pptx / _bench 下 pdf
  · 讲义   = _dump_讲义*.txt / 实时 .docx（裁判，决定 A/B/C）

输出 audit_inputs.json，并打印摘要。**run_audit 会先让用户确认再跑。**

用法: python3 locate_inputs.py --case <dir> [-o audit_inputs.json]
"""
import argparse
import json
import re
from pathlib import Path

# 目录名/文件关键词 → 目标生成 skill
SUBJECT_SKILL = [
    (("英语", "yingyu", "english"), "gaozhong-yingyu-handout-to-courseware", "高中英语"),
    (("化学", "huaxue", "chem"),    "gaozhong-huaxue-handout-to-courseware", "高中化学"),
    (("语文", "古诗", "文言", "yuwen", "gushi"), "yuwen-handout-to-courseware", "高中语文"),
    (("数学", "分解", "牛吃草", "递推", "必胜", "math"), "handout-to-courseware", "小学数学"),
]

EXCLUDE_PPTX = ("标杆", "模版", "模板", "template", "benchmark", "bench")


def _ver(name: str) -> int:
    m = re.search(r"v(\d+)", name, re.I)
    return int(m.group(1)) if m else 0


def _dump_origin(txt_path: Path):
    """读 dump 首行 `# <原始路径>`，回报原始文件路径（可能已不在本机）。"""
    try:
        first = txt_path.read_text(encoding="utf-8", errors="ignore").splitlines()[0]
        if first.startswith("# "):
            return first[2:].strip()
    except Exception:
        pass
    return None


def _pick(globs, root):
    out = []
    for g in globs:
        out += sorted(root.glob(g))
    # 去重保序
    seen, uniq = set(), []
    for p in out:
        if p not in seen:
            seen.add(p)
            uniq.append(p)
    return uniq


def locate(case_dir: str) -> dict:
    root = Path(case_dir).resolve()
    if not root.is_dir():
        raise SystemExit(f"❌ 不是目录: {root}")

    res = {"case_dir": str(root), "case_name": root.name}

    # ── 课件：所有 pptx，排除标杆/模版关键词 ──
    all_pptx = sorted(root.glob("*.pptx"))
    cw = [p for p in all_pptx
          if not any(k in p.name.lower() or k in p.name for k in EXCLUDE_PPTX)]
    cw.sort(key=lambda p: (_ver(p.name), p.stat().st_mtime), reverse=True)
    res["courseware"] = {
        "latest": str(cw[0]) if cw else None,
        "latest_version": _ver(cw[0].name) if cw else None,
        "all_versions": [{"file": str(p), "version": _ver(p.name)} for p in cw],
    }

    # ── 模版 ──
    tspec = root / "template_spec.json"
    tdump = _pick(["_dump_模板*.txt", "_dump_模版*.txt"], root)
    tlive = [p for p in all_pptx if any(k in p.name for k in ("模板", "模版", "template"))]
    res["template"] = {
        "spec_json": str(tspec) if tspec.exists() else None,
        "dump": str(tdump[0]) if tdump else None,
        "live": str(tlive[0]) if tlive else None,
        "origin": _dump_origin(tdump[0]) if tdump else None,
    }

    # ── 标杆 ──
    bdump = _pick(["_dump_标杆*.txt"], root)
    blive = [p for p in all_pptx if "标杆" in p.name]
    bpdf = _pick(["_bench/*.pdf", "*标杆*.pdf"], root)
    res["benchmark"] = {
        "dump": str(bdump[0]) if bdump else None,
        "live": str(blive[0]) if blive else None,
        "pdf": str(bpdf[0]) if bpdf else None,
        "origin": _dump_origin(bdump[0]) if bdump else None,
        "present": bool(bdump or blive or bpdf),
    }

    # ── 讲义 ──
    hdump = _pick(["_dump_讲义*.txt"], root)
    hlive = sorted(root.glob("*.docx"))
    res["handout"] = {
        "dump": str(hdump[0]) if hdump else None,
        "live": str(hlive[0]) if hlive else None,
        "origin": _dump_origin(hdump[0]) if hdump else None,
        "present": bool(hdump or hlive),
    }

    # ── 推断目标生成 skill ──
    hay = (root.name + " " + " ".join(p.name for p in all_pptx)).lower()
    skill = subject = None
    for keys, sk, subj in SUBJECT_SKILL:
        if any(k.lower() in hay for k in keys):
            skill, subject = sk, subj
            break
    # 旁证：生成产物脚本
    gen_artifacts = [p.name for p in root.glob("*.py")
                     if re.search(r"build|render|gen_|extract", p.name, re.I)]
    res["target_skill"] = {"skill": skill, "subject": subject,
                           "gen_artifacts": gen_artifacts,
                           "content_json": str(root / "content.json")
                           if (root / "content.json").exists() else None}

    # ── 缺口告警 ──
    warns = []
    if not res["courseware"]["latest"]:
        warns.append("未找到课件 pptx（对比对象缺失）")
    if not res["benchmark"]["present"]:
        warns.append("未找到标杆（结构/视觉对比将退化为仅完整性+样式）")
    if not res["handout"]["present"]:
        warns.append("未找到讲义（无法做 A/B/C 归因与完整性审计——只能机械列差异）")
    if not (res["template"]["spec_json"] or res["template"]["dump"] or res["template"]["live"]):
        warns.append("未找到模版（样式基准缺失，run_audit 会尝试现拆解 live 模版）")
    res["warnings"] = warns
    return res


def print_summary(res):
    print(f"\n{'='*64}\n输入定位：{res['case_name']}\n{'='*64}")
    cw = res["courseware"]
    print(f"课件   : {Path(cw['latest']).name if cw['latest'] else '—'}"
          f"  (v{cw['latest_version']}，共 {len(cw['all_versions'])} 个版本)")
    if len(cw["all_versions"]) > 1:
        print("         全部版本: " + ", ".join(Path(v["file"]).name for v in cw["all_versions"]))
    t = res["template"]
    print(f"模版   : spec={Path(t['spec_json']).name if t['spec_json'] else '—'}"
          f"  dump={Path(t['dump']).name if t['dump'] else '—'}"
          f"  live={Path(t['live']).name if t['live'] else '—'}")
    b = res["benchmark"]
    print(f"标杆   : dump={Path(b['dump']).name if b['dump'] else '—'}"
          f"  live={Path(b['live']).name if b['live'] else '—'}"
          f"  pdf={Path(b['pdf']).name if b['pdf'] else '—'}")
    h = res["handout"]
    print(f"讲义   : dump={Path(h['dump']).name if h['dump'] else '—'}"
          f"  live={Path(h['live']).name if h['live'] else '—'}")
    ts = res["target_skill"]
    print(f"目标skill: {ts['skill'] or '?'}  ({ts['subject'] or '学科未识别'})")
    if res["warnings"]:
        print("\n⚠️  缺口：")
        for w in res["warnings"]:
            print(f"   - {w}")
    print(f"{'='*64}\n请确认以上对比对象正确后，再跑 run_audit.py（多版本时确认对比哪一版）。")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", required=True, help="case 目录")
    ap.add_argument("-o", "--out", default=None)
    a = ap.parse_args()
    res = locate(a.case)
    out = Path(a.out) if a.out else Path(a.case) / "audit_inputs.json"
    out.write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
    print_summary(res)
    print(f"\n✅ audit_inputs.json → {out}")


if __name__ == "__main__":
    main()
