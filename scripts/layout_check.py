# -*- coding: utf-8 -*-
"""
layout_check.py — 机械检测「排版错位」：文字溢出文本框 / 形状越出画布。

为什么需要它：LibreOffice 渲染会**替换缺失的自定义字体**，从而**掩盖**由真实字体
（普惠体/字魂/汉仪等，往往比替换字更宽更高）撑出文本框造成的溢出——而那正是
WPS/PowerPoint 里看到的「错位」。本检测基于**字体度量估算**，不依赖渲染，能抓到被掩盖的溢出。

估算法（保守）：
  每个文本框 needed_height ≈ Σ_段落 ceil(段落折算宽度 / 每行可容字符) × 字号 × 行距
  CJK 字符宽≈字号(全角)，拉丁/数字宽≈0.55×字号。every paragraph 至少 1 行。
  needed_height > 框高 × 容差 → 溢出嫌疑（给出比值）。
  另：shape 边界超出画布（left<0 / top<0 / right>画布宽 / bottom>画布高）→ 越界。

⚠️ 字号继承：python-pptx 取不到继承字号时按 --default-size（默认18pt）估，结果为**嫌疑**非定论，
   需结合截图/在 WPS 打开确认。

用法: python3 layout_check.py <课件.pptx> [-o layout_issues.json] [--default-size 18] [--tol 1.08]
"""
import argparse
import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu, Pt

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _pt(emu):
    return Emu(emu).pt if emu is not None else None


def _para_lines(text, font_pt, usable_w_pt):
    """估算段落占几行（按 CJK 全角 / 拉丁半角折算宽度）。"""
    if not text.strip():
        return 0
    cjk = len(re.findall(r"[一-鿿　-〿＀-￯]", text))
    other = len(text) - cjk
    width = cjk * font_pt + other * font_pt * 0.55
    per_line = max(1.0, usable_w_pt / max(font_pt, 1))
    width_in_fullwidth = width / max(font_pt, 1)
    return max(1, int(width_in_fullwidth / per_line + 0.999))


def _line_spacing(p):
    """段落行距倍数：spcPct/100000；缺省 1.2。"""
    pPr = p._p.find(f"{A}pPr")
    if pPr is not None:
        ln = pPr.find(f"{A}lnSpc")
        if ln is not None:
            pct = ln.find(f"{A}spcPct")
            if pct is not None:
                return int(pct.get("val", "120000")) / 100000.0
    return 1.2


def check(pptx_path, default_size=18, tol=1.08):
    prs = Presentation(str(pptx_path))
    slide_w, slide_h = _pt(prs.slide_width), _pt(prs.slide_height)
    overflow, offcanvas = [], []

    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.left is None:
                continue
            L, T = _pt(sh.left), _pt(sh.top)
            W, H = _pt(sh.width), _pt(sh.height)
            if W is None or H is None:
                continue
            # 越界（留 1pt 容差）
            over = []
            if L < -1: over.append(f"left={L:.0f}")
            if T < -1: over.append(f"top={T:.0f}")
            if L + W > slide_w + 1: over.append(f"right={L+W:.0f}>{slide_w:.0f}")
            if T + H > slide_h + 1: over.append(f"bottom={T+H:.0f}>{slide_h:.0f}")
            if over:
                offcanvas.append({"slide": i, "shape": sh.name,
                                  "box": [round(L), round(T), round(W), round(H)],
                                  "exceed": over})
            # 文字溢出
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            tf = sh.text_frame
            # autofit=shrink 的框 PowerPoint 会自动缩字，溢出风险低 → 标注但不报
            usable_w = max(W - 6, 6)   # 减去左右内边距估值
            needed = 0.0
            for p in tf.paragraphs:
                runs = [r for r in p.runs if r.text]
                txt = "".join(r.text for r in runs) or p.text
                if not txt.strip():
                    needed += default_size * 0.6   # 空行
                    continue
                sizes = [r.font.size.pt for r in runs if r.font.size]
                fpt = max(sizes) if sizes else default_size
                lines = _para_lines(txt, fpt, usable_w)
                needed += lines * fpt * _line_spacing(p)
            ratio = needed / max(H, 1)
            if ratio > tol:
                overflow.append({"slide": i, "shape": sh.name,
                                 "box_h_pt": round(H), "needed_pt": round(needed),
                                 "ratio": round(ratio, 2),
                                 "text_excerpt": tf.text.strip().replace("\n", " ")[:60]})

    overflow.sort(key=lambda x: -x["ratio"])
    by_slide = {}
    for o in overflow:
        by_slide.setdefault(o["slide"], 0)
        by_slide[o["slide"]] += 1
    return {
        "courseware": Path(pptx_path).name,
        "slide_size_pt": [round(slide_w), round(slide_h)],
        "summary": {"overflow_shapes": len(overflow), "offcanvas_shapes": len(offcanvas),
                    "slides_with_overflow": sorted(by_slide), "params": {"default_size": default_size, "tol": tol}},
        "overflow": overflow,
        "offcanvas": offcanvas,
        "note": "字号继承取不到时按 default_size 估 → 为嫌疑非定论；LibreOffice 渲染会掩盖字体溢出，"
                "本机械检测可补足。逐条结合 WPS/PowerPoint 真渲染确认。",
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--default-size", type=float, default=18)
    ap.add_argument("--tol", type=float, default=1.08)
    a = ap.parse_args()
    res = check(a.pptx, a.default_size, a.tol)
    if a.out:
        Path(a.out).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    s = res["summary"]
    print(f"排版检测: {s['overflow_shapes']} 处文字溢出嫌疑 / {s['offcanvas_shapes']} 处越界")
    print(f"  溢出页: {s['slides_with_overflow']}")
    print("  最严重 15 处:")
    for o in res["overflow"][:15]:
        print(f"    第{o['slide']:>2}页 ratio={o['ratio']} 需{o['needed_pt']}pt/框{o['box_h_pt']}pt 「{o['text_excerpt'][:38]}」")
    if res["offcanvas"]:
        print("  越界:")
        for o in res["offcanvas"][:10]:
            print(f"    第{o['slide']}页 {o['shape']}: {o['exceed']}")
    if a.out:
        print(f"✅ {a.out}")


if __name__ == "__main__":
    main()
