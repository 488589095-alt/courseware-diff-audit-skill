# -*- coding: utf-8 -*-
"""
structure_reconcile.py — 课件 vs 标杆 的结构/页序/页数对账。

产出（机械事实，归因留 Claude）：
  · role_histogram   两边各自的页型计数（按版式名 + 文本关键词粗分）
  · page_binding     标杆每页 → 最相似课件页 + 相似度（低相似度交 Claude 看并排图）
  · benchmark_only   标杆有、课件无的页 → 按「5 类老师增量桶」初分（B 类候选）
  · courseware_only  课件有、标杆无的页（C 类新增候选 / 或 skill 噪声）
  · page_count       total_delta + 各桶计数 + 未归类残差（残差大→报警，Claude 复核）
  · order_milestones 封面/目录/分隔/结束 等里程碑相对顺序

用法:
  python3 structure_reconcile.py --courseware <课件.pptx|dump.txt> \
                                 --benchmark  <标杆.pptx|dump.txt> -o structure_diff.json
"""
import argparse
import json
import re
from pathlib import Path

import _common as C

BIND_THRESHOLD = 0.55

# 页型角色（顺序敏感，先匹配先得）。匹配版式名 + 该页文本。
ROLE_PATTERNS = [
    ("封面",   r"封面|cover|首页"),
    ("目录",   r"目录|contents|content\b"),
    ("分隔",   r"分隔|part\s*\d|章节|篇章|模块.*分隔"),
    ("测评",   r"课前测|课后测|入门测|出门测|预习测|检测|测一测"),
    ("方法讲解", r"步骤|流程|step|方法|技巧|何为|怎么写|如何"),
    ("选择题", r"多选|单选|选择题|选项"),
    ("题源",   r"题源|出处|真题|卷】|·浙江|高考"),
    ("语篇",   r"语篇|passage|原文|阅读材料|短文"),
    ("翻译",   r"翻译|译文|中英|对照"),
    ("练习例题", r"练习|例题|必做|选做|针对训练|实战"),
    ("答案",   r"答案|参考范文|参考答案"),
    ("结束",   r"结束|再见|拜拜|谢谢|thank|the\s*end"),
]

# 5 类老师增量桶（标杆有/课件无时归类）。role → bucket，外加文本旁证。
ROLE_TO_BUCKET = {
    "测评": "测评外壳",
    "方法讲解": "方法讲解",
    "选择题": "互动题+干扰项",
    "翻译": "视觉装饰/过渡/翻译",
}
DECOR_HINT = re.compile(r"你会怎么表达|过渡|青铜|黄金|水平|插画|吉祥物|★|star|加油|恭喜")
INFER_HINT = re.compile(r"推断|整体方向|后续走向|预测|推理(?!.*留白)")
OPTION_PAT = re.compile(r"[A-DＡ-Ｄ][.、．]\s*\S")


def classify_role(layout: str, text: str) -> str:
    hay = f"{layout}  {text[:200]}"
    if len(OPTION_PAT.findall(text)) >= 2:
        return "选择题"
    for role, pat in ROLE_PATTERNS:
        if re.search(pat, hay, re.I):
            return role
    return "正文/其它"


def to_bucket(role: str, text: str) -> str:
    if role in ROLE_TO_BUCKET:
        return ROLE_TO_BUCKET[role]
    if DECOR_HINT.search(text):
        return "视觉装饰/过渡/翻译"
    if INFER_HINT.search(text):
        return "推断填充"
    return "未归类"


def page_summary(s) -> str:
    t = (s.get("text") or "").strip()
    return (t[:36] or f"[{s['layout']}]").replace("\n", " ")


def build_index(deck):
    return [{"no": s["no"], "norm": C.norm(s.get("text", "")),
             "tok": C.tokens(s.get("text", ""))} for s in deck["slides"]]


def best_match(text, index):
    na, ta = C.norm(text), C.tokens(text)
    if not na or not ta:
        return 0.0, None
    cheap = []
    for s in index:
        sub = 1.0 if (na in s["norm"] and len(na) > 8) else 0.0
        ov = len(ta & s["tok"]) / len(ta) if ta else 0.0
        cheap.append((max(sub, ov), s["no"], s["norm"]))
    cheap.sort(key=lambda x: x[0], reverse=True)
    best_sim, best_no = 0.0, None
    for base, no, nrm in cheap[:2]:
        lcs = C._longest_common_substr_ratio(na, nrm)
        sim = max(base, lcs)
        if sim > best_sim:
            best_sim, best_no = sim, no
    return round(best_sim, 3), best_no


def histogram(deck):
    h = {}
    for s in deck["slides"]:
        r = classify_role(s["layout"], s.get("text", ""))
        h[r] = h.get(r, 0) + 1
    return h


def order_milestones(deck):
    layouts = [(s["no"], f"{s['layout']}  {s.get('text','')[:40]}") for s in deck["slides"]]
    out = []
    for name, pat in [("封面", r"封面|cover"), ("目录", r"目录|contents"),
                      ("分隔", r"分隔|part\s*1"), ("结束", r"结束|再见|拜拜|thank")]:
        idx = next((no for no, h in layouts if re.search(pat, h, re.I)), None)
        out.append({"name": name, "page": idx})
    return out


def run(courseware_path, benchmark_path):
    cw = C.load_deck(courseware_path)
    bm = C.load_deck(benchmark_path)
    cw_idx, bm_idx = build_index(cw), build_index(bm)

    # 标杆每页 → 最相似课件页
    page_binding, bench_only = [], []
    buckets = {}
    for s in bm["slides"]:
        sim, cw_no = best_match(s.get("text", ""), cw_idx)
        role = classify_role(s["layout"], s.get("text", ""))
        bound = sim >= BIND_THRESHOLD and cw_no is not None
        page_binding.append({"benchmark_page": s["no"], "benchmark_summary": page_summary(s),
                             "role": role, "courseware_page": cw_no if bound else None,
                             "similarity": sim})
        if not bound:
            bucket = to_bucket(role, s.get("text", ""))
            bench_only.append({"benchmark_page": s["no"], "role": role,
                               "bucket": bucket, "summary": page_summary(s)})
            buckets.setdefault(bucket, []).append(s["no"])

    # 课件每页是否在标杆有对应（找 C 类新增 / 噪声）
    cw_only = []
    for s in cw["slides"]:
        sim, _ = best_match(s.get("text", ""), bm_idx)
        if sim < BIND_THRESHOLD:
            cw_only.append({"courseware_page": s["no"],
                            "role": classify_role(s["layout"], s.get("text", "")),
                            "summary": page_summary(s), "best_benchmark_sim": sim})

    total_delta = bm["total"] - cw["total"]
    bucket_rows = [{"bucket": b, "pages": sorted(p), "count": len(p)}
                   for b, p in sorted(buckets.items(), key=lambda x: -len(x[1]))]
    unbucketed = len(buckets.get("未归类", []))

    # 页型差（主信号）：每页唯一角色 → Σrole_delta 恒等于 total_delta，不依赖脆弱绑定。
    hcw, hbm = histogram(cw), histogram(bm)
    role_delta = []
    for role in sorted(set(hcw) | set(hbm)):
        b, c = hbm.get(role, 0), hcw.get(role, 0)
        if b != c:
            role_delta.append({"role": role, "benchmark": b, "courseware": c, "delta": b - c})
    role_delta.sort(key=lambda r: -r["delta"])
    pos = sum(r["delta"] for r in role_delta if r["delta"] > 0)
    neg = sum(r["delta"] for r in role_delta if r["delta"] < 0)

    return {
        "courseware": {"file": cw["file"], "total": cw["total"], "source": cw["source"]},
        "benchmark": {"file": bm["file"], "total": bm["total"], "source": bm["source"]},
        "role_histogram": {"courseware": hcw, "benchmark": hbm},
        "page_count": {
            "total_delta": total_delta,
            "role_delta": role_delta,                 # 主信号，Σdelta == total_delta
            "benchmark_more_pages": pos, "courseware_more_pages": neg,
            "benchmark_only_by_binding": len(bench_only),
            "courseware_only_by_binding": len(cw_only),
            "unbucketed_benchmark_pages": unbucketed,
            # 提醒：标杆增量页常与课件共享语篇文本→相似度绑定会低估增量；以 role_delta 为准。
            "note": "页数差额以 role_delta 为准(Σ==total_delta)；binding 仅供并排配图，"
                    "标杆增量页常因共享语篇被误绑，benchmark_only_by_binding 会偏小",
            "warn_unbucketed": unbucketed > max(3, len(bench_only) * 0.4),
        },
        "benchmark_only_buckets": bucket_rows,
        "benchmark_only_pages": bench_only,
        "courseware_only_pages": cw_only,
        "page_binding": page_binding,
        "order_milestones": {"courseware": order_milestones(cw),
                             "benchmark": order_milestones(bm)},
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--courseware", required=True)
    ap.add_argument("--benchmark", required=True)
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    res = run(a.courseware, a.benchmark)
    Path(a.out).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    pc = res["page_count"]
    print(f"结构: 标杆 {res['benchmark']['total']} − 课件 {res['courseware']['total']} "
          f"= {pc['total_delta']} 页差额  (标杆多 {pc['benchmark_more_pages']} / 课件多 {pc['courseware_more_pages']})")
    print("  页型差 role_delta (主信号, Σ==total_delta):")
    for r in res["page_count"]["role_delta"]:
        print(f"    {r['role']:<8} 标杆{r['benchmark']:>3} 课件{r['courseware']:>3}  Δ{r['delta']:+d}")
    print("  标杆独有页·增量桶 (B类候选, Claude复核):")
    for b in res["benchmark_only_buckets"]:
        print(f"    [{b['bucket']}] {b['count']} 页  {b['pages'][:12]}")
    if pc["warn_unbucketed"]:
        print("  ⚠️ 未归类增量页偏多，Claude 需看截图细分")
    print(f"✅ {a.out}")


if __name__ == "__main__":
    main()
