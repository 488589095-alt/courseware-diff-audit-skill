# -*- coding: utf-8 -*-
"""
visual_compose.py — 把课件/标杆截图拼成「可一眼对比」的图。

产出两类（对齐手工产物 _grid/ 与 _cmp/）：
  · _grid/课件_sheetN.png / 标杆_sheetN.png  —— 全页缩略网格，一眼看页数差与结构
  · _cmp/pBBB_vs_cCCC.png                     —— 关键页 标杆|课件 并排（带页码标签）

标签用 ASCII（BENCH/CW + 页码），避免缺 CJK 字体时乱码。

可独立调用，但通常由 run_audit.py 传入 pairs（来自 page_binding + completeness 缺失页）。

用法:
  python3 visual_compose.py --cw-png <课件截图目录> --bm-png <标杆截图目录> --out <case目录> \
      [--pairs '[[4,1],[37,18]]']   # [[标杆页, 课件页], ...]
"""
import argparse
import json
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def _font(size):
    for p in ("/System/Library/Fonts/Supplemental/Arial.ttf",
              "/System/Library/Fonts/Helvetica.ttc",
              "/Library/Fonts/Arial.ttf"):
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _pngs(d):
    d = Path(d)
    out = []
    for p in d.glob("slide_*.png"):
        m = re.search(r"slide_(\d+)\.png", p.name)
        if m:
            out.append((int(m.group(1)), p))
    return [p for _, p in sorted(out)]


def _png_for(d, page_no):
    p = Path(d) / f"slide_{page_no:03d}.png"
    return p if p.exists() else None


def make_grid(png_dir, out_path, title, cols=6, thumb_w=260, per_sheet_rows=5):
    pngs = _pngs(png_dir)
    if not pngs:
        return []
    f = _font(15)
    ft = _font(22)
    lbl_h = 20
    per_sheet = cols * per_sheet_rows
    sheets = []
    out_base = Path(out_path)
    out_base.parent.mkdir(parents=True, exist_ok=True)

    for sheet_i in range((len(pngs) + per_sheet - 1) // per_sheet):
        chunk = pngs[sheet_i * per_sheet:(sheet_i + 1) * per_sheet]
        # 用首图比例定缩略高
        with Image.open(chunk[0]) as im0:
            ratio = im0.height / im0.width
        thumb_h = int(thumb_w * ratio)
        cell_w, cell_h = thumb_w + 8, thumb_h + lbl_h + 8
        rows = (len(chunk) + cols - 1) // cols
        title_h = 34
        canvas = Image.new("RGB", (cols * cell_w, title_h + rows * cell_h), "white")
        d = ImageDraw.Draw(canvas)
        d.text((8, 6), f"{title}  (sheet {sheet_i + 1})", fill="black", font=ft)
        for k, png in enumerate(chunk):
            r, c = divmod(k, cols)
            x, y = c * cell_w + 4, title_h + r * cell_h + 4
            page_no = int(re.search(r"slide_(\d+)", png.name).group(1))
            try:
                with Image.open(png) as im:
                    im = im.convert("RGB").resize((thumb_w, thumb_h))
                    canvas.paste(im, (x, y + lbl_h))
            except Exception:
                pass
            d.rectangle([x, y, x + thumb_w, y + lbl_h - 2], fill=(240, 240, 240))
            d.text((x + 4, y + 2), f"p{page_no}", fill=(180, 0, 0), font=f)
        sheet_path = out_base.parent / f"{out_base.stem}_sheet{sheet_i + 1}.png"
        canvas.save(sheet_path)
        sheets.append(str(sheet_path))
    return sheets


def make_pair(bm_png_dir, cw_png_dir, bench_no, cw_no, out_dir, target_w=700):
    bm = _png_for(bm_png_dir, bench_no) if bench_no else None
    cw = _png_for(cw_png_dir, cw_no) if cw_no else None
    if not bm and not cw:
        return None
    f = _font(20)
    gap, head = 16, 30

    def load(p, label):
        if not p:
            ph = Image.new("RGB", (target_w, int(target_w * 0.6)), (245, 245, 245))
            ImageDraw.Draw(ph).text((10, 10), label + " (缺)", fill=(150, 0, 0), font=f)
            return ph
        with Image.open(p) as im:
            im = im.convert("RGB")
            h = int(target_w * im.height / im.width)
            return im.resize((target_w, h))

    left = load(bm, f"BENCH p{bench_no}")
    right = load(cw, f"CW p{cw_no}")
    H = max(left.height, right.height) + head
    canvas = Image.new("RGB", (target_w * 2 + gap, H), "white")
    d = ImageDraw.Draw(canvas)
    d.text((8, 4), f"BENCH p{bench_no}", fill=(0, 0, 160), font=f)
    d.text((target_w + gap + 8, 4), f"CW p{cw_no}", fill=(0, 130, 0), font=f)
    canvas.paste(left, (0, head))
    canvas.paste(right, (target_w + gap, head))
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"p{(bench_no or 0):03d}_vs_c{(cw_no or 0):03d}.png"
    canvas.save(out_path)
    return str(out_path)


def run(cw_png_dir, bm_png_dir, out_root, pairs=None):
    out_root = Path(out_root)
    grids = {}
    grids["courseware"] = make_grid(cw_png_dir, out_root / "_grid" / "课件.png", "COURSEWARE")
    if bm_png_dir and _pngs(bm_png_dir):
        grids["benchmark"] = make_grid(bm_png_dir, out_root / "_grid" / "标杆.png", "BENCHMARK")
    cmp_imgs = []
    if pairs and bm_png_dir:
        for bench_no, cw_no in pairs:
            p = make_pair(bm_png_dir, cw_png_dir, bench_no, cw_no, out_root / "_cmp")
            if p:
                cmp_imgs.append(p)
    return {"grids": grids, "pairs": cmp_imgs}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cw-png", required=True)
    ap.add_argument("--bm-png", default=None)
    ap.add_argument("--out", required=True)
    ap.add_argument("--pairs", default=None, help="JSON: [[bench_no, cw_no], ...]")
    a = ap.parse_args()
    pairs = json.loads(a.pairs) if a.pairs else None
    res = run(a.cw_png, a.bm_png, a.out, pairs)
    print("缩略网格:")
    for k, sheets in res["grids"].items():
        for s in sheets:
            print(f"  {k}: {s}")
    if res["pairs"]:
        print("并排对比:")
        for p in res["pairs"]:
            print(f"  {p}")


if __name__ == "__main__":
    main()
