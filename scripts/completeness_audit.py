# -*- coding: utf-8 -*-
"""
completeness_audit.py — 以讲义为裁判，逐项核对讲义内容是否已转入课件（A 类完整性）。

脚本只判机械事实：每条讲义内容在课件里 hit / partial / miss + 命中页 + 相似度。
**绝不判断「该不该有」**——是 A 类 bug 还是 B 类老师增量，由 Claude 按
classification_framework.md 结合 §note 与并排截图归因。

讲义切块（subject-agnostic）：
  · 【答案】【参考答案】【参考范文】 → type=答案（必须转入，漏=高危）
  · 【知识标签】                    → type=知识标签（预期落在备注 notes）
  · 【解析】【点拨】【方法】         → type=解析（常为标杆/老师专属，低优先）
  · 【注】                          → type=注释（属内容，需转入）
  · 表格                            → type=表格（每张表一条）
  · Heading 段落                    → type=章节标题
  · 留白模板（“1基础版本：____”）   → type=留白（学生填，内容在【答案】里，不计入分母）
  · 其它实质正文                    → type=正文

用法:
  python3 completeness_audit.py --handout <讲义.docx|_dump_讲义.txt> \
                                --courseware <课件.pptx|_dump_课件.txt> -o completeness.json
"""
import argparse
import json
import re
from pathlib import Path

import _common as C

ANSWER = re.compile(r"^[\s\d.、]*【(答案|参考答案|参考范文|范文)】")
TAG = re.compile(r"^[\s\d.、]*【(知识标签|标签)】")
ANALYSIS = re.compile(r"^[\s\d.、]*【(解析|点拨|方法|技巧|思路)】")
NOTE = re.compile(r"^[\s\d.、]*【注】")
BLANK_TPL = re.compile(r"_{6,}")          # 6+ 下划线 = 留白模板
PREFIX_STRIP = re.compile(r"^[\s\d.、]*【[^】]*】\s*[:：]?\s*")
ZHEN_PREFIX = re.compile(r"^(中文|英文|译文|参考)\s*[:：]")
IMG_HINT = re.compile(r"(knowledge\s*map|思维导图|知识图谱|流程图|示意图|结构图|配图|image\d|图\d)", re.I)


def _clean_match_text(t: str) -> str:
    """去掉 markdown 式标记/前缀，留下要匹配的内容主体。"""
    t = PREFIX_STRIP.sub("", t)
    t = ZHEN_PREFIX.sub("", t).strip()
    return t


def segment_handout(doc) -> list:
    """把讲义切成可审计条目。返回 [{id, handout_ref, section, type, text_excerpt,
       match_text, is_image, priority}]。"""
    items = []
    section = ""
    n = 0

    def add(ref, typ, raw, match_text, priority, is_image=False):
        nonlocal n
        n += 1
        items.append({
            "id": f"H-{n:03d}", "handout_ref": ref, "section": section, "type": typ,
            "text_excerpt": raw[:120], "match_text": match_text,
            "is_image": is_image, "priority": priority,
        })

    for it in doc["items"]:
        txt = it["text"].strip()
        if not txt:
            continue
        ref = f"P{it['idx']}"
        style = (it.get("style") or "")
        is_heading = "heading" in style.lower() or "标题" in style or "title" in style.lower()

        if ANSWER.match(txt):
            add(ref, "答案", txt, _clean_match_text(txt), "high")
        elif TAG.match(txt):
            add(ref, "知识标签", txt, _clean_match_text(txt), "notes")
        elif ANALYSIS.match(txt):
            body = _clean_match_text(txt)
            add(ref, "解析", txt, body, "low")
        elif NOTE.match(txt):
            add(ref, "注释", txt, _clean_match_text(txt), "high")
        elif is_heading:
            add(ref, "章节标题", txt, txt, "medium", is_image=bool(IMG_HINT.search(txt)))
        elif BLANK_TPL.search(txt) and len(C.norm(BLANK_TPL.sub("", txt))) < 8:
            add(ref, "留白", txt, "", "skip")           # 学生填空模板，内容在【答案】
        elif len(C.norm(txt)) >= 6:
            add(ref, "正文", txt, _clean_match_text(txt), "medium",
                is_image=bool(IMG_HINT.search(txt)))
        # 过短碎片忽略

        if is_heading:
            section = txt[:40]

    for t in doc["tables"]:
        cells = [c for row in t["rows"] for c in row if c.strip()]
        blob = " ".join(cells)
        if C.norm(blob):
            n += 1
            items.append({
                "id": f"H-{n:03d}", "handout_ref": f"TABLE {t['idx']}", "section": section,
                "type": "表格", "text_excerpt": blob[:120], "match_text": blob,
                "is_image": False, "priority": "high",
            })
    return items


def build_slide_index(deck):
    """每页可搜文本 = 正文 + 表格 + 备注；预归一化 + token 集，供两段式匹配。"""
    idx = []
    for s in deck["slides"]:
        blob = (s.get("text") or "") + "  " + (s.get("notes") or "")
        idx.append({"no": s["no"], "norm": C.norm(blob), "tok": C.tokens(blob)})
    return idx


def best_matches(match_text, slide_idx, hit=0.8, part=0.5):
    """两段式：先 token 重叠 + 子串筛 top-3，再对 top-3 跑 LCS 取最大。"""
    na = C.norm(match_text)
    ta = C.tokens(match_text)
    if not na or not ta:
        return {"status": "skip", "courseware_pages": [], "similarity": 0.0, "matched_excerpt": ""}
    cheap = []
    for s in slide_idx:
        sub = 1.0 if na in s["norm"] else 0.0
        ov = len(ta & s["tok"]) / len(ta) if ta else 0.0
        cheap.append((max(sub, ov), s))
    cheap.sort(key=lambda x: x[0], reverse=True)

    scored = []
    for base, s in cheap[:3]:
        lcs = C._longest_common_substr_ratio(na, s["norm"])
        scored.append((max(base, lcs), s["no"]))
    scored.sort(reverse=True)
    best = scored[0][0] if scored else 0.0
    # 命中页：cheap 阶段 >= hit 阈值的所有页（覆盖被拆到多页的情况）
    hit_pages = sorted({s["no"] for c, s in cheap if c >= hit})
    if best >= hit:
        pages = hit_pages or [scored[0][1]]
        status = "hit"
    elif best >= part:
        pages = [scored[0][1]]
        status = "partial"
    else:
        pages = []
        status = "miss"
    return {"status": status, "courseware_pages": pages,
            "similarity": round(best, 3), "matched_excerpt": ""}


def run(handout_path, courseware_path):
    doc = C.load_handout(handout_path)
    deck = C.load_deck(courseware_path)
    slide_idx = build_slide_index(deck)
    items = segment_handout(doc)

    audited = []
    for it in items:
        if it["priority"] == "skip" or it["type"] == "留白":
            it["match"] = {"status": "skip", "courseware_pages": [], "similarity": 0.0}
            audited.append(it)
            continue
        m = best_matches(it["match_text"], slide_idx)
        if it["is_image"] and m["status"] == "miss":
            m["status"] = "image_check"      # 图片类无法文本匹配，交目检
        it["match"] = m
        it.pop("match_text", None)
        audited.append(it)

    # 可审计 = 有内容可匹配的条目（排除留白模板与空体标记如空【解析】）
    graded = [i for i in audited if i["priority"] != "skip" and i["type"] != "留白"
              and i["match"]["status"] != "skip"]
    cnt = lambda st: sum(1 for i in graded if i["match"]["status"] == st)
    total = len(graded)
    hit, part, miss, imgc = cnt("hit"), cnt("partial"), cnt("miss"), cnt("image_check")
    coverage = round((hit + 0.5 * part) / total, 3) if total else 1.0

    # 按 type 分层——headline 数字会被「短标题弱匹配」「留白表格」拉低，
    # 真实信号看 答案/注释/正文 的 hit 率（这几类 miss 才是 A 类 bug 强信号）。
    by_type = {}
    for i in graded:
        st = i["match"]["status"]
        d = by_type.setdefault(i["type"], {"hit": 0, "partial": 0, "miss": 0, "image_check": 0})
        d[st] = d.get(st, 0) + 1
    content_types = ("答案", "注释", "正文", "表格")
    ct = [i for i in graded if i["type"] in content_types]
    content_cov = round(
        (sum(1 for i in ct if i["match"]["status"] == "hit")
         + 0.5 * sum(1 for i in ct if i["match"]["status"] == "partial")) / len(ct), 3
    ) if ct else 1.0

    return {
        "handout": doc["file"], "handout_source": doc["source"],
        "courseware": deck["file"], "courseware_source": deck["source"],
        "items": audited,
        "summary": {"total_graded": total, "hit": hit, "partial": part,
                    "miss": miss, "image_check": imgc, "coverage": coverage,
                    "content_coverage": content_cov, "by_type": by_type,
                    "skipped_blanks": sum(1 for i in audited if i["type"] == "留白"),
                    "note": "miss/image_check 才需判 A/B/C；partial 多为短标题弱匹配或留白表格(设计如此)，看 by_type 与截图确认"},
        "miss_items": [{"id": i["id"], "ref": i["handout_ref"], "type": i["type"],
                        "section": i["section"], "excerpt": i["text_excerpt"]}
                       for i in graded if i["match"]["status"] in ("miss", "image_check")],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--handout", required=True)
    ap.add_argument("--courseware", required=True)
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    res = run(a.handout, a.courseware)
    Path(a.out).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    s = res["summary"]
    print(f"完整性: {s['hit']}/{s['total_graded']} hit, {s['partial']} partial, "
          f"{s['miss']} miss, {s['image_check']} 待目检 → coverage {s['coverage']} "
          f"(内容类 content_coverage {s['content_coverage']})")
    print("  分层: " + "  ".join(f"{t}={d.get('hit',0)}/{sum(d.values())}h"
                                  for t, d in s["by_type"].items()))
    if res["miss_items"]:
        print("疑似缺失（交 Claude 判 A/B/C）:")
        for m in res["miss_items"][:25]:
            print(f"  ✗ {m['id']} [{m['type']}] {m['ref']} «{m['excerpt'][:50]}»")
    print(f"✅ {a.out}")


if __name__ == "__main__":
    main()
